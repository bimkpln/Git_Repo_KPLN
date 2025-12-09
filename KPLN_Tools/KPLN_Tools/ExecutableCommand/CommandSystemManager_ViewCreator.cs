using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using KPLN_Tools.ExternalCommands;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExecutableCommand
{
    internal class CommandSystemManager_ViewCreator : IExecutableCommand
    {
        private readonly OVVK_SystemManager_VM _viewModel;
        private readonly string[] _systemSumParameters;
        private readonly List<ElementId> _docCatIdColl = new List<ElementId>();

        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _errorDict = new Dictionary<string, List<ElementId>>();
        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _warningDict = new Dictionary<string, List<ElementId>>();

        public CommandSystemManager_ViewCreator(OVVK_SystemManager_VM currentViewModel, string[] systemSumParameters)
        {
            _viewModel = currentViewModel;
            _systemSumParameters = systemSumParameters;
        }

        public Result Execute(UIApplication app)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName($"{Command_OVVK_SystemManager.PluginName}_Создание видов", ModuleData.ModuleName).ConfigureAwait(false);

            #region Подготовка коллекции и элементов
            Autodesk.Revit.DB.View activeView = _viewModel.CurrentDoc.ActiveView;

            _docCatIdColl.AddRange(_viewModel.ElementColl.Select(e => e.Category.Id).Distinct());

            // Вывод критических ошибок и остановка работы
            if (_errorDict.Keys.Count != 0)
            {
                MessageBox.Show(
                   $"В модели есть критические ошибки. Работа прервана, список выведен отдельным окном",
                   "Внимание",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Asterisk);
                PrintErrors(_errorDict);

                return Result.Cancelled;
            }

            // Получение шаблона вида
            Autodesk.Revit.DB.View viewTempl = (Autodesk.Revit.DB.View)_viewModel.CurrentDoc.GetElement(activeView.get_Parameter(BuiltInParameter.VIEW_TEMPLATE).AsElementId());
            if (viewTempl == null)
            {
                DialogResult userRes = MessageBox.Show(
                   $"У данного вида НЕТ шаблона. Такой подход не рекомендуется использовать, " +
                    $"т.к. последующая поддержка изменений будет затруднена (нужно будет создавать шаблон и присваивать его отдельно всем видам). " +
                    $"Все равно продолжить?",
                   "Внимание",
                   MessageBoxButtons.YesNo,
                   MessageBoxIcon.Asterisk);

                if (userRes != DialogResult.Yes)
                    return Result.Cancelled;
            }

            // Анализ парамтеров шаблона на блокировку парамтеров фильтрации
            if (viewTempl != null)
            {
                ICollection<ElementId> viewTemplNonControlIds = viewTempl.GetNonControlledTemplateParameterIds();
                bool isFiltersNonControlled = viewTemplNonControlIds.Any();
                foreach (ElementId id in viewTemplNonControlIds)
                {
                    // Для ревит 2020 нет возможности вытянуть чере API элемент фильтров шаблона. НО ID всегда один и тот же, даже для Ревит 2023 (но лучше отдельно прописать)
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                    if (id.IntegerValue == -1006964)
#else
                    if (id.Value == -1006964)
#endif
                    {
                        isFiltersNonControlled = true;
                        break;
                    }
                    else
                        isFiltersNonControlled = false;
                }

                if (!isFiltersNonControlled)
                {
                    MessageBox.Show(
                        $"В шаблоне вида галка, управляющая фильтрами должна быть СНЯТА. Работа экстренно завершена!",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return Result.Cancelled;
                }
            }

            #endregion

            #region Обработка коллекции
            using (Transaction t = new Transaction(_viewModel.CurrentDoc, $"KPLN: Создание видов"))
            {
                t.Start();

                if (!CreateViews(_viewModel.CurrentDoc, viewTempl))
                {
                    t.RollBack();
                    return Result.Cancelled;
                }
                else
                    t.Commit();
            }

            if (_warningDict.Keys.Count != 0)
            {
                PrintErrors(_warningDict);
                MessageBox.Show(
                    $"Скрипт отработал, но есть замечания к некоторым элементам. Список выведен отдельным окном",
                    "Внимание",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show(
                    $"Скрипт отработал без ошибок",
                    "Успешно",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            #endregion

            return Result.Succeeded;
        }

        /// <summary>
        /// Создать виды (аксонометрические схемы)
        /// </summary>
        private bool CreateViews(Document doc, Autodesk.Revit.DB.View viewTempl)
        {
            Dictionary<ElementId, string> docViewDict = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .Where(v => !v.IsTemplate)
                .ToDictionary(d => d.Id, d => d.Name);

            ParameterFilterElement[] docFilters = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .ToArray();

            foreach (string sysName in _systemSumParameters)
            {
                string filterName = $"prog_{_viewModel.ParameterName} = НРВ_{sysName}";

#if Revit2020 || Debug2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(_viewModel.ElementColl.FirstOrDefault().LookupParameter(_viewModel.ParameterName).Id, sysName, false);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(_viewModel.ElementColl.FirstOrDefault().LookupParameter(_viewModel.ParameterName).Id, sysName);
#endif
                ElementParameterFilter filterRules = new ElementParameterFilter(fRule);

                ElementId viewFilterId = null;
                ParameterFilterElement oldEqualFRule = docFilters.FirstOrDefault(df => df.Name.Equals(filterName));
                if (oldEqualFRule != null)
                    viewFilterId = oldEqualFRule.Id;
                else
                {
                    ParameterFilterElement newViewFilter = ParameterFilterElement.Create(doc, filterName, _docCatIdColl, filterRules);
                    viewFilterId = newViewFilter.Id;
                }

                string newViewName = sysName.Contains(_viewModel.SysNameSeparator) ? $"Схема систем_{sysName}" : $"Схема системы_{sysName}";
                bool viewNameError = false;
                foreach (KeyValuePair<ElementId, string> kvp in docViewDict)
                {
                    if (kvp.Value.Equals(newViewName))
                    {
                        AddToErrorDict(_warningDict, $"Вид с имением {newViewName} уже существует", kvp.Key);
                        viewNameError = true;
                    }
                }

                if (viewNameError)
                    continue;

                Autodesk.Revit.DB.View newView = doc.GetElement(doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate)) as Autodesk.Revit.DB.View;
                newView.AddFilter(viewFilterId);
                newView.SetFilterVisibility(viewFilterId, false);

                if (viewTempl != null)
                    newView.ApplyViewTemplateParameters(viewTempl);

                newView.Name = newViewName;
            }

            return true;
        }

        /// <summary>
        /// Добавить эл-т ревит к замечанию
        /// </summary>
        /// <param name="errMsg"></param>
        /// <param name="elementId"></param>
        private void AddToErrorDict(Dictionary<string, List<ElementId>> dict, string errMsg, ElementId elementId)
        {
            if (dict.ContainsKey(errMsg))
            {
                List<ElementId> ids = dict[errMsg];
                ids.Add(elementId);

                dict[errMsg] = ids;
            }
            else
                dict.Add(errMsg, new List<ElementId> { elementId });
        }

        /// <summary>
        /// Вывод ошибок пользователю
        /// </summary>
        private void PrintErrors(Dictionary<string, List<ElementId>> dict)
        {
            foreach (KeyValuePair<string, List<ElementId>> kvp in dict)
            {
                string errorElemsId = string.Join(", ", kvp.Value);
                HtmlOutput.Print($"ОШИБКА: \"{kvp.Key}\" у элементов: {errorElemsId}", MessageType.Error);
            }
        }
    }
}
