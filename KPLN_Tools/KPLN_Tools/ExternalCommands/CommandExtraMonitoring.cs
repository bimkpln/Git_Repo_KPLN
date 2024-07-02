using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandExtraMonitoring : IExternalCommand
    {
        private readonly Dictionary<ElementId, List<MonitorEntity>> _monitorEntitiesDict = new Dictionary<ElementId, List<MonitorEntity>>();
        private readonly Dictionary<ElementId, List<MonitorLinkEntity>> _monitorLinkEntiteDict = new Dictionary<ElementId, List<MonitorLinkEntity>>();
        private RevitLinkInstance _currentLink;

        /// <summary>
        /// Коллекция ошибок, при отработке
        /// </summary>
        private readonly Dictionary<string, List<Element>> _localErrors = new Dictionary<string, List<Element>>();

        /// <summary>
        /// Коллекция категорий, котоыре могут мониториться в проектах КПЛН
        /// </summary>
        internal BuiltInCategory[] MonitoredBuiltInCatArr { get; } = new BuiltInCategory[]
        {
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctTerminal,
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            ElementId[] selectedIds = uidoc.Selection.GetElementIds().ToArray();
            if (selectedIds.Length == 0)
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                    MainContent = "Предварительно необходимо выбрать элементы, котоыре нужно проанализировать, а потом - запустить плагин",
                }
                .Show();

                return Result.Failed;
            }
            
            try
            {
                SetMonitoredElemsFromUserSelect(doc, selectedIds);

                if (_monitorEntitiesDict.Count > 0)
                {
                    List<ElementId> errorElemIdColl = new List<ElementId>();
                    foreach (var id in selectedIds)
                    {
                        bool inColl = false;
                        foreach (var kvp in _monitorEntitiesDict)
                        {
                            foreach (var v in kvp.Value)
                            {
                                if (v.ModelElement != null && v.ModelElement.Id.IntegerValue == id.IntegerValue)
                                {
                                    inColl = true;
                                    break;
                                }
                            }
                        }
                        if (!inColl)
                            errorElemIdColl.Add(id);
                    }

                    if (errorElemIdColl.Count > 0)
                    {
                        string msg = string.Join(", ", errorElemIdColl);
                        Print(
                            $"Элементы из выборки - не удалось определить основу. Нужно чтобы элементы у связи и твоего проекта - полностью совпадали по геометрии. Ошибки: {msg}",
                            MessageType.Error);

                        Print("Работа экстренно остановлена. Исправь ошибки: ", MessageType.Error);

                        return Result.Failed;
                    }
                }

                if (_monitorEntitiesDict.Count > 0)
                {
                    MonitoringParamSetter setterForm = new MonitoringParamSetter(doc, _monitorEntitiesDict);
                    setterForm.ShowDialog();
                }

                if (_localErrors.Count > 0)
                {
                    foreach (KeyValuePair<string, List<Element>> kvp in _localErrors)
                    {
                        string errorIds = String.Empty;
                        foreach (Element element in kvp.Value)
                        {
                            errorIds = string.Join(", ", element.Id);
                        }
                    
                        Print($"{kvp.Key} {errorIds}", MessageType.Error);
                    }
                    Print($"В ходе работы были выявлены ошибки:", MessageType.Error);
                }
            }
            catch (Exception ex)
            {
                new TaskDialog("Ошибка")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainContent = ex.Message,
                }
                .Show();

                return Result.Failed;
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получить коллекцию элементов из связи для пользовательской выборки
        /// </summary>
        /// <param name="commandData"></param>
        /// <returns></returns>
        private void SetMonitoredElemsFromUserSelect(Document doc, ElementId[] selectedIds)
        {
            // Очистка выборки от случайных элементов
            Element[] trueElems = selectedIds
                .Where(id => MonitoredBuiltInCatArr.Contains((BuiltInCategory)doc.GetElement(id).Category.Id.IntegerValue))
                .Select(id => doc.GetElement(id))
                .ToArray();
            if (trueElems.Length == 0)
                throw new Exception("Ни один выбранный элемент не попадает под функцию мониторинга");

            #region Подготовка элементов, которые нужно проверить
            // Подготовка точек
            HashSet<XYZ> locPntsFromUserSelection = new HashSet<XYZ>();
            foreach (Element elem in trueElems)
            {
                if (elem.Location is LocationPoint locPoint)
                    locPntsFromUserSelection.Add(locPoint.Point);
                else
                    ErrorDictSetting("Элементы из выборки - не удалось получить LocationPoint. Скинь в BIM-отдел:", elem);
            }

            // Подготовка параметров элементов проекта
            HashSet<Parameter> docElemsParams = GetParametersFromElems(trueElems);
            #endregion

            #region Подготовка BoundingBoxXYZ для анализа связи
            double maxX = locPntsFromUserSelection.Max(pnt => pnt.X);
            double maxY = locPntsFromUserSelection.Max(pnt => pnt.Y);
            double maxZ = locPntsFromUserSelection.Max(pnt => pnt.Z);
            XYZ maxPoint = new XYZ(maxX, maxY, maxZ);

            double minX = locPntsFromUserSelection.Min(pnt => pnt.X);
            double minY = locPntsFromUserSelection.Min(pnt => pnt.Y);
            double minZ = locPntsFromUserSelection.Min(pnt => pnt.Z);
            XYZ minPoint = new XYZ(minX, minY, minZ);

            // Расширяем для малой выборки
            if (minPoint.IsAlmostEqualTo(maxPoint, 1))
            {
                maxPoint += new XYZ(1.0, 1.0, 1.0);
                minPoint -= new XYZ(1.0, 1.0, 1.0);
            }

            BoundingBoxXYZ searchBbox = new BoundingBoxXYZ() { Max = maxPoint, Min = minPoint };
            #endregion

            #region Обработка элементов
            foreach (Element element in trueElems)
            {
                if (element != null)
                {
                    ElementId[] monitoredLinkElemIdsArr = element.GetMonitoredLinkElementIds().ToArray();
                    if (monitoredLinkElemIdsArr.Length == 0)
                        throw new Exception($"Элемент с id:{element.Id} - не имеет мониторинга!");
                    else if (monitoredLinkElemIdsArr.Length > 1)
                        throw new Exception($"Элемент с id:{element.Id} - имеет мониторинг из нескольких связей. Это запрещено!");
                    else
                    {
                        RevitLinkInstance currentLink = doc.GetElement(monitoredLinkElemIdsArr[0]) as RevitLinkInstance;
                        if (_currentLink == null)
                        {
                            _currentLink = currentLink;
                            PreapareMonitorEntityColl(searchBbox);
                        }
                        else if (_currentLink.Id.IntegerValue != currentLink.Id.IntegerValue)
                            throw new Exception($"Работа экстренно прекращена! Элемент с id:{element.Id} - имеет мониторинг из другой связи. Можно выполнить проверку только с разделением по связям.");
                    
                        UpdateMonitorEntityColl(doc, element, docElemsParams);
                    }
                }
                else
                    throw new Exception($"Работа экстренно прекращена! Скинь в BIM-отдел: элемент с id:{element.Id} - невозможно преобразовать к базоваму классу Element");
            }
            #endregion
        }

        private void ErrorDictSetting(string errorMsg, Element element)
        {
            if (_localErrors.ContainsKey(errorMsg))
                _localErrors[errorMsg].Add(element);
            else
                _localErrors[errorMsg] = new List<Element>() { element };
        }

        /// <summary>
        /// Получить коллекцию параметров, котоыре есть в каждом выделенном элементе
        /// </summary>
        /// <param name="trueElems"></param>
        /// <returns></returns>
        private HashSet<Parameter> GetParametersFromElems(Element[] trueElems)
        {
            HashSet<string> docElemsParamNames = new HashSet<string>();
            docElemsParamNames.UnionWith(trueElems.SelectMany(elem => elem.GetOrderedParameters().Select(p => p.Definition.Name)));
            docElemsParamNames.UnionWith(trueElems.SelectMany(elem => (elem as FamilyInstance).Symbol.GetOrderedParameters().Select(p => p.Definition.Name)));

            HashSet<Parameter> result = new HashSet<Parameter>();
            foreach (string paramName in docElemsParamNames)
            {
                Parameter param = null;
                bool isContain = true;
                foreach (Element elem in trueElems)
                {
                    param = elem.LookupParameter(paramName);
                    if (param == null)
                        param = (elem as FamilyInstance).Symbol.LookupParameter(paramName);
                    
                    if (param == null)
                    {
                        isContain = false;
                        break;
                    }
                }

                if (isContain)
                    result.Add(param);
            }

            return result;
        }

        /// <summary>
        /// Подготовить коллекцию спец. элементов
        /// Когда-нибудь - добавят в Revit API возможность забрать элемент из связи, а пока - работаем с геометрией.
        /// https://forums.autodesk.com/t5/revit-ideas/copy-monitor-api/idi-p/6322737
        /// </summary>
        /// <returns></returns>
        [Obsolete]
        private void PreapareMonitorEntityColl(BoundingBoxXYZ searchBbox)
        {
            if (!_monitorLinkEntiteDict.ContainsKey(_currentLink.Id))
                _monitorLinkEntiteDict[_currentLink.Id] = new List<MonitorLinkEntity>();

            Transform linkTrans = _currentLink.GetTransform();
            foreach (BuiltInCategory bic in MonitoredBuiltInCatArr)
            {
                // Добавяляю элементы из связи
                IList<Element> bicElems = new FilteredElementCollector(_currentLink.GetLinkDocument())
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements();

                List<Element> intersectedBicElems = new List<Element>();
                foreach (Element el in bicElems)
                {
                    if (el.Location is LocationPoint locPoint)
                    {
                        XYZ elPntTransformed = linkTrans.OfPoint(locPoint.Point);
                        if (elPntTransformed.X >= searchBbox.Min.X && elPntTransformed.X <= searchBbox.Max.X &&
                            elPntTransformed.Y >= searchBbox.Min.Y && elPntTransformed.Y <= searchBbox.Max.Y &&
                            elPntTransformed.Z >= searchBbox.Min.Z && elPntTransformed.Z <= searchBbox.Max.Z)
                        {
                            intersectedBicElems.Add(el);
                        }
                    }
                    else
                        ErrorDictSetting($"Элементы из связи {_currentLink.Name} - не удалось получить LocationPoint. Скинь в BIM-отдел:", el);
                }

                // Добавляю параметры из связи
                HashSet<Parameter> linkElemsParams = GetParametersFromElems(intersectedBicElems.ToArray());

                foreach (Element el in intersectedBicElems)
                {
                    _monitorLinkEntiteDict[_currentLink.Id].Add(new MonitorLinkEntity(el, linkElemsParams, _currentLink));
                }
            }

            if (_monitorLinkEntiteDict[_currentLink.Id].Count == 0)
                Print($"Элементы из связи {_currentLink.Name} - не удалось получить элементы проекта. Скинь в BIM-отдел!", MessageType.Error);
        }

        /// <summary>
        /// Обновить коллекцию спец. элементов по пересечению с элементом из проекта
        /// Когда-нибудь - добавят в Revit API возможность забрать элемент из связи, а пока - работаем с геометрией.
        /// https://forums.autodesk.com/t5/revit-ideas/copy-monitor-api/idi-p/6322737
        /// </summary>
        /// <returns></returns>
        [Obsolete]
        private void UpdateMonitorEntityColl(Document doc, Element modelElement, HashSet<Parameter> docElemsParams)
        {
            Solid modelElemSolid = MonitorTool.GetSolidFromElem(modelElement);

            foreach (KeyValuePair<ElementId, List<MonitorLinkEntity>> kvp in _monitorLinkEntiteDict)
            {
                if (!_monitorEntitiesDict.ContainsKey(_currentLink.Id))
                    _monitorEntitiesDict[_currentLink.Id] = new List<MonitorEntity>();

                if (kvp.Key.IntegerValue == _currentLink.Id.IntegerValue)
                {
                    RevitLinkInstance linkInstance = doc.GetElement(kvp.Key) as RevitLinkInstance;
                    Document linkDoc = linkInstance.GetLinkDocument();

                    foreach (MonitorLinkEntity monitorLinkEntity in kvp.Value)
                    {
                        Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            modelElemSolid,
                            monitorLinkEntity.LinkElementSolid, 
                            BooleanOperationsType.Intersect);
                        if (intersectionSolid != null && intersectionSolid.Volume > 0)
                        {
                            _monitorEntitiesDict[_currentLink.Id].Add(
                                new MonitorEntity(
                                    modelElement,
                                    modelElemSolid,
                                    docElemsParams,
                                    monitorLinkEntity
                                    ));

                            return;
                        }
                    }
                }
            }
        }
    }
}
