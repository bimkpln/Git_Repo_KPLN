using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;
using static KPLN_Loader.Preferences;


namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandExtraMonitoring : IExternalCommand
    {
        private readonly Dictionary<ElementId, List<MonitorEntity>> _monitorEntitiesDict = new Dictionary<ElementId, List<MonitorEntity>>();
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
            ElementId[] trueElemIds = selectedIds
                .Where(id => MonitoredBuiltInCatArr.Contains((BuiltInCategory)doc.GetElement(id).Category.Id.IntegerValue))
                .ToArray();
            if (trueElemIds.Length == 0)
                throw new Exception("Ни один выбранный элемент не попадает под функцию мониторинга");

            #region Подготовка элементов, которые нужно проверить
            // Подготовка точек
            HashSet<XYZ> locPntsFromUserSelection = new HashSet<XYZ>();
            foreach (ElementId id in trueElemIds)
            {
                Element elem = doc.GetElement(id);
                if (elem.Location is LocationPoint locPoint)
                    locPntsFromUserSelection.Add(locPoint.Point);
                else
                    ErrorDictSetting("Элементы из выборки - не удалось получить LocationPoint. Скинь в BIM-отдел:", elem);
            }

            // Подготовка параметров элементов проекта
            HashSet<Parameter> docElemsParams = GetParametersFromElems(doc, trueElemIds);
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

            BoundingBoxXYZ searchBbox = new BoundingBoxXYZ() { Max = maxPoint, Min = minPoint };
            #endregion

            #region Обработка элементов
            foreach (ElementId id in trueElemIds)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    ElementId[] monitoredLinkElemIdsArr = element.GetMonitoredLinkElementIds().ToArray();
                    if (monitoredLinkElemIdsArr.Length == 0)
                        throw new Exception($"Элемент с id:{id} - не имеет мониторинга!");
                    else if (monitoredLinkElemIdsArr.Length > 1)
                        throw new Exception($"Элемент с id:{id} - имеет мониторинг из нескольких связей. Это запрещено!");
                    else
                    {
                        RevitLinkInstance currentLink = doc.GetElement(monitoredLinkElemIdsArr[0]) as RevitLinkInstance;
                        if (_currentLink == null)
                        {
                            _currentLink = currentLink;
                            PreapareMonitorEntityColl(_currentLink, searchBbox);
                        }
                        else if (_currentLink.Id.IntegerValue != currentLink.Id.IntegerValue)
                            throw new Exception($"Работа экстренно прекращена! Элемент с id:{id} - имеет мониторинг из другой связи. Можно выполнить проверку только с разделением по связям.");

                        UpdateMonitorEntityColl(currentLink, element, docElemsParams);
                    }
                }
                else
                    throw new Exception($"Работа экстренно прекращена! Скинь в BIM-отдел: элемент с id:{id} - невозможно преобразовать к базоваму классу Element");
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
        /// <param name="doc"></param>
        /// <param name="trueElemIds"></param>
        /// <returns></returns>
        private HashSet<Parameter> GetParametersFromElems(Document doc, ElementId[] trueElemIds)
        {
            HashSet<ElementId> docElemsParamIds = new HashSet<ElementId>();
            docElemsParamIds.UnionWith(trueElemIds.SelectMany(id => doc.GetElement(id).GetOrderedParameters().Select(p => p.Id)));
            docElemsParamIds.UnionWith(trueElemIds.SelectMany(id => (doc.GetElement(id) as FamilyInstance).Symbol.GetOrderedParameters().Select(p => p.Id)));

            HashSet<Parameter> result = new HashSet<Parameter>();
            foreach (ElementId paramId in docElemsParamIds)
            {
                Parameter param = null;
                bool isContain = true;
                foreach (ElementId elId in trueElemIds)
                {
                    param = doc.GetElement(elId).get_Parameter((BuiltInParameter)paramId.IntegerValue);
                    if (param == null)
                        param = (doc.GetElement(elId) as FamilyInstance).Symbol.get_Parameter((BuiltInParameter)paramId.IntegerValue);

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
        private void PreapareMonitorEntityColl(RevitLinkInstance currentLink, BoundingBoxXYZ searchBbox)
        {
            if (!_monitorEntitiesDict.ContainsKey(currentLink.Id))
                _monitorEntitiesDict[currentLink.Id] = new List<MonitorEntity>();

            Transform linkTrans = currentLink.GetTransform();
            foreach (BuiltInCategory bic in MonitoredBuiltInCatArr)
            {
                IList<Element> bicElems = new FilteredElementCollector(currentLink.GetLinkDocument())
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements();
                foreach (Element el in bicElems)
                {
                    if (el.Location is LocationPoint locPoint)
                    {
                        XYZ elPntTransformed = linkTrans.OfPoint(locPoint.Point);
                        if (elPntTransformed.X >= searchBbox.Min.X && elPntTransformed.X <= searchBbox.Max.X &&
                            elPntTransformed.Y >= searchBbox.Min.Y && elPntTransformed.Y <= searchBbox.Max.Y &&
                            elPntTransformed.Z >= searchBbox.Min.Z && elPntTransformed.Z <= searchBbox.Max.Z)
                        {
                            _monitorEntitiesDict[currentLink.Id].Add(new MonitorEntity(el, currentLink));
                        }
                    }
                    else
                        ErrorDictSetting($"Элементы из связи {currentLink.Name} - не удалось получить LocationPoint. Скинь в BIM-отдел:", el);
                }
            }

            // Добавляю параметры
            HashSet<Parameter> linkElemsParams = GetParametersFromElems(
                currentLink.GetLinkDocument(),
                _monitorEntitiesDict[currentLink.Id].Select(m => m.LinkElement.Id).ToArray());
            foreach (var kvp in _monitorEntitiesDict)
            {
                foreach (var monitorEntity in kvp.Value)
                {
                    monitorEntity.LinkElemsParams = linkElemsParams;
                }
            }

            if (_monitorEntitiesDict[currentLink.Id].Count == 0)
                Print($"Элементы из связи {currentLink.Name} - не удалось получить элементы проекта. Скинь в BIM-отдел!", MessageType.Error);
        }

        /// <summary>
        /// Обновить коллекцию спец. элементов по пересечению с элементом из проекта
        /// Когда-нибудь - добавят в Revit API возможность забрать элемент из связи, а пока - работаем с геометрией.
        /// https://forums.autodesk.com/t5/revit-ideas/copy-monitor-api/idi-p/6322737
        /// </summary>
        /// <returns></returns>
        [Obsolete]
        private void UpdateMonitorEntityColl(RevitLinkInstance currentLink, Element modelElement, HashSet<Parameter> docElemsParams)
        {
            foreach (KeyValuePair<ElementId, List<MonitorEntity>> kvp in _monitorEntitiesDict)
            {
                if (kvp.Key.IntegerValue == currentLink.Id.IntegerValue)
                {
                    foreach (MonitorEntity monitorEntity in kvp.Value)
                    {
                        Solid modelElemSolid = MonitorEntity.GetSolidFromElem(modelElement);
                        Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            modelElemSolid,
                            monitorEntity.LinkElementSolid,
                            BooleanOperationsType.Intersect);
                        if (intersectionSolid != null && intersectionSolid.Volume > 0)
                        {
                            monitorEntity.ModelElement = modelElement;
                            monitorEntity.ModelElementSolid = modelElemSolid;
                            monitorEntity.ModelParameters = docElemsParams;
                        }
                    }
                }
            }
        }
    }
}