using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.ExternalCommands;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class ExcCmd_SystemManager_ViewCreator : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly string[] _systemSumParameters;
        private readonly bool _dataIsCurrent;
        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _warningDict = new Dictionary<string, List<ElementId>>();

        public ExcCmd_SystemManager_ViewCreator(OVVK_SystemManagerM currentViewModel, string[] systemSumParameters,
            bool dataIsCurrent = false)
        {
            _model = currentViewModel;
            _systemSumParameters = systemSumParameters;
            _dataIsCurrent = dataIsCurrent;
        }

        public Result Execute(UIApplication app)
        {
            if (!_dataIsCurrent)
                _model.Refresh();
            if (_model.HasParameterErrors || _model.HasSystemValidationErrors || _model.HasDuplicateSystemNames)
            {
                _model.SetOperationMessage("Создание 3D-видов недоступно: исправь все критические ошибки, отмеченные красным.");
                return Result.Cancelled;
            }

            #region Подготовка коллекции и элементов
            Autodesk.Revit.DB.View activeView = _model.CurrentDoc.ActiveView;

            if (!(activeView is View3D) || activeView.IsTemplate
                || !activeView.CanViewBeDuplicated(ViewDuplicateOption.Duplicate))
            {
                _model.SetOperationMessage("Открой 3D-вид, на основе которого нужно создать схемы систем.");
                return Result.Cancelled;
            }

            // Получение шаблона вида
            Autodesk.Revit.DB.View viewTempl = _model.CurrentDoc.GetElement(activeView.ViewTemplateId) as Autodesk.Revit.DB.View;
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

            // Повторная проверка перед изменением модели.
            if (!ValidateTemplateFilters(viewTempl))
                return Result.Cancelled;

            #endregion

            #region Обработка коллекции
            using (Transaction t = new Transaction(_model.CurrentDoc, $"KPLN: Создание видов"))
            {
                if (t.Start() != TransactionStatus.Started)
                    return Result.Cancelled;

                if (!CreateViews(_model.CurrentDoc, viewTempl))
                {
                    t.RollBack();
                    return Result.Cancelled;
                }
                if (t.Commit() != TransactionStatus.Committed)
                {
                    _model.SetOperationMessage("Создание 3D-видов отменено: Revit не подтвердил изменения.");
                    return Result.Cancelled;
                }
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

        internal static bool ValidateTemplateFilters(Autodesk.Revit.DB.View viewTemplate)
        {
            if (viewTemplate == null)
                return true;

            // В Revit 2020 параметр управления фильтрами проверяется по его постоянному ID.
            bool filtersNotControlled = viewTemplate.GetNonControlledTemplateParameterIds().Any(id =>
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                id.IntegerValue == -1006964);
#else
                id.Value == -1006964);
#endif
            if (filtersNotControlled)
                return true;

            MessageBox.Show(
                "В шаблоне вида галка, управляющая фильтрами должна быть СНЯТА. Работа экстренно завершена!",
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return false;
        }

        /// <summary>
        /// Создать виды (аксонометрические схемы)
        /// </summary>
        private bool CreateViews(Document doc, Autodesk.Revit.DB.View viewTempl)
        {
            Dictionary<string, ElementId> docViews;
            using (var collector = new FilteredElementCollector(doc))
                docViews = collector.OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>().Where(view => !view.IsTemplate)
                    .GroupBy(view => view.Name, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Id, StringComparer.OrdinalIgnoreCase);

            Dictionary<string, ParameterFilterElement> docFilters;
            using (var collector = new FilteredElementCollector(doc))
                docFilters = collector.OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .ToDictionary(filter => filter.Name, StringComparer.OrdinalIgnoreCase);

            string filterNamePrefix = "prog_" + _model.ParameterName + " = НРВ_";
            var systemViewFilterIds = new HashSet<ElementId>(docFilters.Values
                .Where(filter => filter.Name.StartsWith(filterNamePrefix, StringComparison.Ordinal))
                .Select(filter => filter.Id));

            // Доступность параметров категории одинакова для всех создаваемых схем.
            var filterableCategories = new HashSet<ElementId>(ParameterFilterUtilities.GetAllFilterableCategories());
            var categoryParameters = _model.ElementColl.Where(element => element.Category != null)
                .Select(element => element.Category.Id).Distinct().Where(filterableCategories.Contains)
                .ToDictionary(id => id, id => new HashSet<ElementId>(
                    ParameterFilterUtilities.GetFilterableParametersInCommon(doc, new[] { id })));
            var excludedIds = new HashSet<ElementId>(_model.WarningsElementColl.Concat(_model.TypeParameterElementColl)
                .Select(item => item.Id));

            foreach (string sysName in _systemSumParameters.Distinct(StringComparer.Ordinal))
            {
                string newViewName = sysName.Contains(_model.SysNameSeparator) ? $"Схема систем_{sysName}" : $"Схема системы_{sysName}";
                if (docViews.TryGetValue(newViewName, out ElementId existingViewId))
                {
                    AddToErrorDict(_warningDict, $"Вид с именем {newViewName} уже существует", existingViewId);
                    continue;
                }

                // Параметр берём у выбранной корректной системы, а не у первого элемента модели.
                Element element = _model.GetSystemElementIds(new OVVK_SystemManagerRowM(sysName))
                    .Select(id => doc.GetElement(id)).FirstOrDefault(item => item != null);
                Parameter parameter = element?.LookupParameter(_model.ParameterName);
                if (parameter == null || parameter.StorageType != StorageType.String)
                    throw new InvalidOperationException("Система «" + sysName + "» больше не доступна. Обнови список систем.");

                List<ElementId> categories = categoryParameters.Where(pair => pair.Value.Contains(parameter.Id))
                    .Select(pair => pair.Key).ToList();
                if (categories.Count == 0)
                    throw new InvalidOperationException("Не найдены категории для фильтра системы «" + sysName + "».");

                string filterName = filterNamePrefix + sysName;

#if Revit2020 || Debug2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(parameter.Id, sysName, false);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(parameter.Id, sysName);
#endif
                ElementParameterFilter filterRules = new ElementParameterFilter(fRule);

                ElementId viewFilterId = null;
                if (docFilters.TryGetValue(filterName, out ParameterFilterElement oldEqualFRule))
                {
                    oldEqualFRule.SetCategories(categories);
                    oldEqualFRule.SetElementFilter(filterRules);
                    viewFilterId = oldEqualFRule.Id;
                }
                else
                {
                    ParameterFilterElement newViewFilter = ParameterFilterElement.Create(doc, filterName, categories, filterRules);
                    viewFilterId = newViewFilter.Id;
                    docFilters.Add(filterName, newViewFilter);
                }

                Autodesk.Revit.DB.View newView = doc.GetElement(doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate)) as Autodesk.Revit.DB.View;
                if (viewTempl != null)
                    newView.ApplyViewTemplateParameters(viewTempl);

                // При копировании готовой схемы фильтр прежней системы не должен скрыть новую.
                foreach (ElementId filterId in newView.GetFilters().Where(systemViewFilterIds.Contains).ToArray())
                    newView.RemoveFilter(filterId);
                if (!newView.GetFilters().Contains(viewFilterId))
                    newView.AddFilter(viewFilterId);
                newView.SetFilterVisibility(viewFilterId, false);

                // Элементы замечаний и категории без этого параметра не входят в схему системы.
                var categoryIds = new HashSet<ElementId>(categories);
                ElementId[] hiddenIds = _model.ElementColl
                    .Where(item => excludedIds.Contains(item.Id) || item.Category == null || !categoryIds.Contains(item.Category.Id))
                    .Where(item => item.CanBeHidden(newView) && !item.IsHidden(newView))
                    .Select(item => item.Id).Distinct().ToArray();
                if (hiddenIds.Length > 0)
                    newView.HideElements(hiddenIds);

                newView.Name = newViewName;
                docViews.Add(newViewName, newView.Id);
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
