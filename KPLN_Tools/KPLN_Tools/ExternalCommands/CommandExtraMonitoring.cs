using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_PluginActivityWorker;
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
        internal const string PluginName = "Экстрамониторинг";

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
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);
                
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
                                if (v.ModelElement != null && v.ModelElement.Id.Equals(id))
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
        private void SetMonitoredElemsFromUserSelect(Document doc, ElementId[] selectedIds)
        {
            // Очистка выборки от случайных элементов
            Element[] trueElems = selectedIds
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                .Where(id => MonitoredBuiltInCatArr.Contains((BuiltInCategory)doc.GetElement(id).Category.Id.IntegerValue))
#else
                .Where(id => MonitoredBuiltInCatArr.Contains(doc.GetElement(id).Category.BuiltInCategory))
#endif
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

            #region Подготовка крайних точек для Outline для анализа связи
            double maxX = locPntsFromUserSelection.Max(pnt => pnt.X);
            double maxY = locPntsFromUserSelection.Max(pnt => pnt.Y);
            double maxZ = locPntsFromUserSelection.Max(pnt => pnt.Z);
            XYZ maxPoint = new XYZ(maxX + 0.5, maxY + 0.5, maxZ + 0.5);

            double minX = locPntsFromUserSelection.Min(pnt => pnt.X);
            double minY = locPntsFromUserSelection.Min(pnt => pnt.Y);
            double minZ = locPntsFromUserSelection.Min(pnt => pnt.Z);
            XYZ minPoint = new XYZ(minX - 0.5, minY - 0.5, minZ - 0.5);
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

                            Transform linkTrans = _currentLink.GetTransform();
                            Outline outlineForFilter = new Outline(linkTrans.Inverse.OfPoint(minPoint), linkTrans.Inverse.OfPoint(maxPoint));
                            // СЛАБОЕ МЕСТО: Если БТП имеют разное положение - при иверсии точек происходит смещение на Basic, и точки могут перевенуться
                            if (outlineForFilter.IsEmpty)
                            {
                                double outMaxX = outlineForFilter.MaximumPoint.X;
                                double outMaxY = outlineForFilter.MaximumPoint.Y;
                                double outMaxZ = outlineForFilter.MaximumPoint.Z;

                                double outMinX = outlineForFilter.MinimumPoint.X;
                                double outMinY = outlineForFilter.MinimumPoint.Y;
                                double outMinZ = outlineForFilter.MinimumPoint.Z;

                                double resOutMaxX = outMaxX > outMinX ? outMaxX : outMinX;
                                double resOutMaxY = outMaxY > outMinY ? outMaxY : outMinY;
                                double resOutMaxZ = outMaxZ > outMinZ ? outMaxZ : outMinZ;

                                double resOutMinX = outMaxX > outMinX ? outMinX : outMaxX;
                                double resOutMinY = outMaxY > outMinY ? outMinY : outMaxY;
                                double resOutMinZ = outMaxZ > outMinZ ? outMinZ : outMaxZ;

                                outlineForFilter = new Outline(new XYZ(resOutMinX, resOutMinY, resOutMinZ), new XYZ(resOutMaxX, resOutMaxY, resOutMaxZ));
                            }
                            PreapareMonitorEntityColl(outlineForFilter);
                        }
                        else if (!_currentLink.Id.Equals(currentLink.Id))
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
        [Obsolete]
        private void PreapareMonitorEntityColl(Outline outlineForFilter)
        {
            if (!_monitorLinkEntiteDict.ContainsKey(_currentLink.Id))
                _monitorLinkEntiteDict[_currentLink.Id] = new List<MonitorLinkEntity>();

            BoundingBoxIntersectsFilter bboxFilter = new BoundingBoxIntersectsFilter(outlineForFilter, 1);
            foreach (BuiltInCategory bic in MonitoredBuiltInCatArr)
            {
                // Добавяляю элементы из связи
                Element[] intersectedBicElems = new FilteredElementCollector(_currentLink.GetLinkDocument())
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .WherePasses(bboxFilter)
                    .ToElements()
                    .ToArray();
                
                // Добавляю параметры из связи
                HashSet<Parameter> linkElemsParams = GetParametersFromElems(intersectedBicElems);

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
