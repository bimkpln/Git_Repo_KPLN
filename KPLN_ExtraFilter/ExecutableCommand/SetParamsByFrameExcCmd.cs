using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.Entities.SetParamsByFrame;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_ExtraFilter.ExecutableCommand
{
    internal sealed class SetParamsByFrameExcCmd : IExecutableCommand
    {
        private readonly SetParamsByFrameM _entity;
        private readonly Dictionary<string, List<ElementId>> _warningsElementColl = new Dictionary<string, List<ElementId>>();
        private readonly HashSet<ElementId> _warningElementIds = new HashSet<ElementId>();

        public SetParamsByFrameExcCmd(SetParamsByFrameM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;

            Document doc = uidoc.Document;

            // Запись конфигурации последнего запуска
            ConfigService.SaveConfig<SetParamsByFrameM_ParamM>(ModuleData.RevitVersion, doc, ConfigType.Memory, _entity.ParamItems);

            _warningsElementColl.Clear();
            _warningElementIds.Clear();

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(SetParamsByFrameExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
            try
            {
                // Запуск рамки самодостаточен и по итогу - завершается 
                if (_entity.CurrentScript == SetParamsByFrameScript.SelectElementsByFrame)
                {
                    IEnumerable<Element> selectedByFrame = SelectionSearchFilter.UserSelectedFilters(uidoc);
                    _entity.UserSelElems = selectedByFrame;

                    if (selectedByFrame != null)
                        _entity.SetRunResult(SetParamsByFrameMessageKind.Success, $"Выбрано элементов: {selectedByFrame.Count()}.");

                    return Result.Succeeded;
                }
                
                
                // Получаю коллекцию экстравыбора
                Element[] extraSel = ExtraSelection(_entity.Doc, _entity.UserSelElems)
                    .Where(e => e != null)
                    .ToArray();

                using (Transaction trans = new Transaction(doc, "KPLN: Выбрать/заполнить рамкой"))
                {
                    trans.Start();

                    foreach (Element elem in extraSel)
                    {
                        foreach (SetParamsByFrameM_ParamM paramM in _entity.ParamItems)
                        {
                            string paramName = paramM.ParamM_SelectedParameter.RevitParamName;
                            string newValue = paramM.ParamM_InputValue;

                            Parameter currentParam = elem.LookupParameter(paramName)
                                ?? elem.Document.GetElement(elem.GetTypeId())?.LookupParameter(paramName);
                            if (currentParam == null)
                            {
                                // Не у всех элементов есть опция "Параметр типа", такие в отчёт не отправляем
                                if (!IsRoomSeparationLine(elem))
                                    AddWarning($"Отсутствует параметр {paramName} (как в типе, так и в экземпляре)", elem.Id);
                                
                                continue;
                            }

                            try
                            {
                                switch (currentParam.StorageType)
                                {
                                    case StorageType.Double:
                                        if (double.TryParse(newValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double dValue))
                                        {
#if Debug2020 || Revit2020
                                            DisplayUnitType unitType = currentParam.DisplayUnitType;
                                            double prjData = UnitUtils.ConvertToInternalUnits(dValue, unitType);
#else
                                            ForgeTypeId forgeTypeId = currentParam.GetUnitTypeId();
                                            double prjData = UnitUtils.ConvertToInternalUnits(dValue, forgeTypeId);
#endif
                                            SetParameterValue(currentParam, prjData, elem.Id, paramName);
                                        }
                                        else
                                            AddWarning($"Для параметра {paramName} был выбран не верный тип ввода (вместо числа - текст)", elem.Id);
                                        break;
                                    case StorageType.Integer:
                                        if (int.TryParse(newValue, out int iValue))
                                            SetParameterValue(currentParam, iValue, elem.Id, paramName);
                                        else
                                            AddWarning($"Для параметра {paramName} был выбран не верный тип ввода (вместо целого числа - текст)", elem.Id);
                                        break;
                                    case StorageType.String:
                                        SetParameterValue(currentParam, newValue, elem.Id, paramName);
                                        break;
                                    default:
                                        AddWarning($"Для параметра {paramName} был выбран не верный тип ввода", elem.Id);
                                        break;
                                }
                            }
                            catch (Exception paramEx)
                            {
                                AddWarning($"Не удалось заполнить параметр {paramName}: {paramEx.Message}", elem.Id);
                            }
                        }
                    }

                    trans.Commit();
                }

                int totalElementCount = extraSel.Length;
                int warningElementCount = _warningElementIds.Count;
                int successElementCount = Math.Max(0, totalElementCount - warningElementCount);

                // Отправляю экстровыбор пользователю
                app.ActiveUIDocument.Selection.SetElementIds(extraSel.Select(e => e.Id).ToList());

                if (warningElementCount == 0)
                    _entity.SetRunResult(SetParamsByFrameMessageKind.Success, $"Обработано элементов: {successElementCount}.");
                else if (successElementCount == 0)
                {
                    _entity.SetRunResult(
                        SetParamsByFrameMessageKind.Error,
                        $"Не обработано ни одного элемента: {successElementCount} из {totalElementCount}. Детальный отчёт будет открыт в отдельном окне.");

                    HtmlOutput.PrintMsgDict("Не выполнено", MessageType.Error, _warningsElementColl);
                }
                else
                {
                    _entity.SetRunResult(
                        SetParamsByFrameMessageKind.Warning,
                        $"Обработано успешно: {successElementCount} из {totalElementCount}. Детальный отчёт будет открыт в отдельном окне.");

                    HtmlOutput.PrintMsgDict("Выполнено с замечаниями", MessageType.Warning, _warningsElementColl);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                _entity.SetRunResult(SetParamsByFrameMessageKind.Error, "Критическая ошибка. Детальный отчёт будет открыт в отдельном окне.");
                HtmlOutput.Print($"Ошибка попытки выбора подобных. Отправь разработчику: {ex.Message}",
                    MessageType.Error);

                return Result.Cancelled;
            }
        }


        private void AddWarning(string message, ElementId elementId)
        {
            HtmlOutput.SetMsgDict_ByMsg(message, elementId, _warningsElementColl);
            _warningElementIds.Add(elementId);
        }

        private void SetParameterValue(Parameter parameter, double value, ElementId elementId, string paramName)
        {
            if (!parameter.Set(value))
                AddWarning($"Не удалось заполнить параметр {paramName}", elementId);
        }

        private void SetParameterValue(Parameter parameter, int value, ElementId elementId, string paramName)
        {
            if (!parameter.Set(value))
                AddWarning($"Не удалось заполнить параметр {paramName}", elementId);
        }

        private void SetParameterValue(Parameter parameter, string value, ElementId elementId, string paramName)
        {
            if (!parameter.Set(value))
                AddWarning($"Не удалось заполнить параметр {paramName}", elementId);
        }


        private static bool IsRoomSeparationLine(Element elem)
        {
            if (elem?.Category == null)
                return false;

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_RoomSeparationLines;
#else
            return elem.Category.BuiltInCategory == BuiltInCategory.OST_RoomSeparationLines;
#endif
        }

        /// <summary>
        /// Расширенное выделение элементов модели
        /// </summary>
        /// <param name="doc">Ревит-док</param>
        /// <param name="selectedElems">Коллекция выделенных в ревит эл-в</param>
        /// <returns></returns>
        private static IEnumerable<Element> ExtraSelection(Document doc, IEnumerable<Element> selectedElems)
        {
            List<Element> result = new List<Element>(selectedElems);

            ElementClassFilter famIsntFilter = new ElementClassFilter(typeof(FamilyInstance));
            ElementClassFilter ductInsulFilter = new ElementClassFilter(typeof(DuctInsulation));
            ElementClassFilter pipeInsulFilter = new ElementClassFilter(typeof(PipeInsulation));

            List<ElementFilter> filters = new List<ElementFilter>() { famIsntFilter, ductInsulFilter, pipeInsulFilter };
            LogicalOrFilter resultFilter = new LogicalOrFilter(filters);

            foreach (Element elem in selectedElems)
            {
                IList<ElementId> depElems = elem.GetDependentElements(resultFilter);
                foreach (ElementId id in depElems)
                {
                    Element currentElem = doc.GetElement(id);
                    if (currentElem.Id.Equals(elem.Id))
                        continue;

                    // Игнорирую балясины (отдельно в спеки не идут)
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                    BuiltInCategory bic = (BuiltInCategory)currentElem.Category.Id.IntegerValue;
#else
                    BuiltInCategory bic = currentElem.Category.BuiltInCategory;
#endif
                    if (bic == BuiltInCategory.OST_StairsRailingBaluster)
                        continue;

                    // Предварительно фильтрую общие вложенные семейства
                    if (currentElem is FamilyInstance famInst && famInst.SuperComponent != null)
                        result.Add(famInst);
                    // Добавляю ВСЕ остальное
                    else
                        result.Add(currentElem);
                }
            }

            return result;
        }
    }
}
