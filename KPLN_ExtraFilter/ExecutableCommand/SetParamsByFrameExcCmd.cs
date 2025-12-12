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
    internal class SetParamsByFrameExcCmd : IExecutableCommand
    {
        private readonly SetParamsByFrameM _entity;
        private readonly Dictionary<string, List<ElementId>> _warningsElementColl = new Dictionary<string, List<ElementId>>();

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

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(SetParamsByFrameExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
            try
            {
                // Запуск рамки самодостаточен и по итогу - завершается 
                if (_entity.CurrentScript == SetParamsByFrameScript.SelectElementsByFrame)
                {
                    _entity.UserSelElems = SelectionSearchFilter.UserSelectedFilters(uidoc);
                    return Result.Succeeded;
                }
                
                
                // Получаю коллекцию экстравыбора
                IEnumerable<Element> extraSel = ExtraSelection(_entity.Doc, _entity.UserSelElems);

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
                                ?? elem.Document.GetElement(elem.GetTypeId()).LookupParameter(paramName);
                            if (currentParam == null)
                            {
                                HtmlOutput.SetMsgDict_ByMsg($"Отсутствует параметр {paramName} (как в типе, так и в экземпляре)", elem.Id, _warningsElementColl);
                                continue;
                            }

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
                                        currentParam.Set(prjData);
                                        break;
                                    }
                                    goto case StorageType.Integer;
                                case StorageType.Integer:
                                    if (int.TryParse(newValue, out int iValue))
                                    {
                                        currentParam.Set(iValue);
                                        break;
                                    }
                                    goto case StorageType.String;
                                case StorageType.String:
                                    currentParam.Set(newValue);
                                    break;
                                default:
                                    HtmlOutput.SetMsgDict_ByMsg($"Для параметра {paramName} был выбран не верный тип ввода (вместо цифры - текст)", elem.Id, _warningsElementColl);
                                    break;
                            }
                        }
                    }

                    trans.Commit();
                }

                // Отправляю экстровыбор пользователю
                app.ActiveUIDocument.Selection.SetElementIds(extraSel.Select(e => e.Id).ToList());

                HtmlOutput.PrintMsgDict("Выполнено с замечаниями", MessageType.Warning, _warningsElementColl);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.Print($"Ошибка попытки выбора подобных. Отправь разработчику: {ex.Message}",
                    MessageType.Error);

                return Result.Cancelled;
            }
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
