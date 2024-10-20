using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
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
        private readonly OVVK_SystemManager_ViewModel _currentViewModel;
        private readonly List<Element> _docElemsColl = new List<Element>();
        private readonly List<ElementId> _docCatIdColl = new List<ElementId>();

        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _errorDict = new Dictionary<string, List<ElementId>>();
        /// <summary>
        /// Словарь элементов, где ключ - текст ошибки, значения - список элементов
        /// </summary>
        private readonly Dictionary<string, List<ElementId>> _warningDict = new Dictionary<string, List<ElementId>>();

        public CommandSystemManager_ViewCreator(OVVK_SystemManager_ViewModel vm)
        {
            _currentViewModel = vm;
        }

        public Result Execute(UIApplication app)
        {
            #region Подготовка коллекции и элементов
            Document doc = _currentViewModel.CurrentDoc;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            // Анализ вида
            if (activeView == null || activeView.ViewType != ViewType.ThreeD)
            {
                MessageBox.Show(
                    $"Скрипт нужно запускать при открытом 3D-виде, т.к. на основании его будут создаваиться аналоги",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return Result.Cancelled;
            }

            // Готовлю коллекцию элементов для анализа
            foreach (BuiltInCategory bic in Command_OVVK_SystemManager.FamileInstanceBICs)
            {
                _docElemsColl.AddRange(new FilteredElementCollector(doc, activeView.Id)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(Command_OVVK_SystemManager.FamilyNameFilter));
            }

            foreach (BuiltInCategory bic in Command_OVVK_SystemManager.MEPCurveBICs)
            {
                _docElemsColl.AddRange(new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType());
            }

            _docCatIdColl.AddRange(_docElemsColl.Select(e => e.Category.Id));

            // Вывод критических ошибок и остановка работы
            if (_errorDict.Keys.Count != 0)
            {
                MessageBox.Show(
                   $"В модели есть критические ошибки. Работа прервана, список выведен отдельным окном",
                   "Внимание",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Asterisk);
                PrintErrors( _errorDict );

                return Result.Cancelled;
            }

            // Получение шаблона вида
            Autodesk.Revit.DB.View viewTempl = (Autodesk.Revit.DB.View)doc.GetElement(activeView.get_Parameter(BuiltInParameter.VIEW_TEMPLATE).AsElementId());
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
                    if (id.IntegerValue == -1006964)
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
            using (Transaction t = new Transaction(doc, $"KPLN: Создание видов"))
            {
                t.Start();

                if (!CreateViews(doc, viewTempl))
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
            foreach (string sysName in _currentViewModel.SystemSumParameters)
            {
                string filterName = $"prog_{_currentViewModel.ParameterName} = !*{sysName}*!";

                FilterRule fRule = ParameterFilterRuleFactory.CreateNotContainsRule(_docElemsColl.FirstOrDefault().LookupParameter(_currentViewModel.ParameterName).Id, sysName, false);
                ElementParameterFilter filterRules = new ElementParameterFilter(fRule);
                ParameterFilterElement newViewFilter = ParameterFilterElement.Create(doc, filterName, _docCatIdColl, filterRules);
                ElementId newViewFilterId = newViewFilter.Id;

                Autodesk.Revit.DB.View newView = doc.GetElement(doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate)) as Autodesk.Revit.DB.View;
                newView.AddFilter(newViewFilterId);
                newView.SetFilterVisibility(newViewFilterId, false);

                if (viewTempl != null)
                    newView.ApplyViewTemplateParameters(viewTempl);

                if (sysName.Contains(_currentViewModel.SysNameSeparator))
                    newView.Name = $"Схема систем_{sysName}";
                else
                    newView.Name = $"Схема системы_{sysName}";
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
