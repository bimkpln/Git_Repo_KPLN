using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.ExternalCommands;
using KPLN_Tools_OVVK.Forms.Entities;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class ExcCmd_SystemManager_ViewCreator : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
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

        public ExcCmd_SystemManager_ViewCreator(OVVK_SystemManagerM currentViewModel, string[] systemSumParameters)
        {
            _model = currentViewModel;
            _systemSumParameters = systemSumParameters;
        }

        public Result Execute(UIApplication app)
        {
            #region Подготовка коллекции и элементов
            Autodesk.Revit.DB.View activeView = _model.CurrentDoc.ActiveView;

            _docCatIdColl.AddRange(_model.ElementColl.Where(e => e.Category != null).Select(e => e.Category.Id).Distinct());

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
            Autodesk.Revit.DB.View viewTempl = (Autodesk.Revit.DB.View)_model.CurrentDoc.GetElement(activeView.get_Parameter(BuiltInParameter.VIEW_TEMPLATE).AsElementId());
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
                t.Start();

                if (!CreateViews(_model.CurrentDoc, viewTempl))
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
                // Параметр берём у выбранной корректной системы, а не у первого элемента модели.
                Element element = _model.GetSystemElementIds(new OVVK_SystemManagerRowM(sysName))
                    .Select(id => doc.GetElement(id)).FirstOrDefault(item => item != null);
                Parameter parameter = element?.LookupParameter(_model.ParameterName);
                if (parameter == null || parameter.StorageType != StorageType.String)
                    continue;

                var filterableCategories = new HashSet<ElementId>(ParameterFilterUtilities.GetAllFilterableCategories());
                List<ElementId> categories = _docCatIdColl.Where(id => filterableCategories.Contains(id)
                    && ParameterFilterUtilities.GetFilterableParametersInCommon(doc, new[] { id }).Contains(parameter.Id)).ToList();
                if (categories.Count == 0)
                    throw new System.InvalidOperationException("Не найдены категории для фильтра системы «" + sysName + "».");

                string filterName = $"prog_{_model.ParameterName} = НРВ_{sysName}";

#if Revit2020 || Debug2020
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(parameter.Id, sysName, false);
#else
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotEqualsRule(parameter.Id, sysName);
#endif
                ElementParameterFilter filterRules = new ElementParameterFilter(fRule);

                ElementId viewFilterId = null;
                ParameterFilterElement oldEqualFRule = docFilters.FirstOrDefault(df => df.Name.Equals(filterName));
                if (oldEqualFRule != null)
                {
                    oldEqualFRule.SetCategories(categories);
                    oldEqualFRule.SetElementFilter(filterRules);
                    viewFilterId = oldEqualFRule.Id;
                }
                else
                {
                    ParameterFilterElement newViewFilter = ParameterFilterElement.Create(doc, filterName, categories, filterRules);
                    viewFilterId = newViewFilter.Id;
                }

                string newViewName = sysName.Contains(_model.SysNameSeparator) ? $"Схема систем_{sysName}" : $"Схема системы_{sysName}";
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
                if (!newView.GetFilters().Contains(viewFilterId))
                    newView.AddFilter(viewFilterId);
                newView.SetFilterVisibility(viewFilterId, false);

                if (viewTempl != null)
                    newView.ApplyViewTemplateParameters(viewTempl);

                // Элементы замечаний и категории без этого параметра не входят в схему системы.
                var excludedIds = new HashSet<ElementId>(_model.WarningsElementColl.Concat(_model.TypeParameterElementColl)
                    .Select(item => item.Id));
                ElementId[] hiddenIds = _model.ElementColl
                    .Where(item => excludedIds.Contains(item.Id) || item.Category == null || !categories.Contains(item.Category.Id))
                    .Where(item => item.CanBeHidden(newView) && !item.IsHidden(newView))
                    .Select(item => item.Id).Distinct().ToArray();
                if (hiddenIds.Length > 0)
                    newView.HideElements(hiddenIds);

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
