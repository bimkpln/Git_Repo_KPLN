using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExecutableCommand
{
    internal class SelectionByClickExcCmd : IExecutableCommand
    {
        private readonly SelectionByClickM _entity;

        /// <summary>
        /// Итоговая коллекция, которую нужно выделить в модели
        /// </summary>
        private readonly List<ElementId> _resultColl = new List<ElementId>();

        public SelectionByClickExcCmd(SelectionByClickM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            Document doc = uiDoc.Document;

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(SelectionByModelExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
            try
            {
                // Поиск элементов комбинированием фильтров
                List<ElementFilter> filters = new List<ElementFilter>();

                #region Определяю коллекцию для анализа
                FilteredElementCollector mainFIC = null;
                if (_entity.Where_Model)
                    mainFIC = new FilteredElementCollector(doc).WhereElementIsNotElementType();
                else if (_entity.Where_CurrentView)
                {
                    View currentView = doc.ActiveView
                        ?? throw new System.Exception("Отправь разработчику: Не удалось определить открытый вид");

                    mainFIC = new FilteredElementCollector(doc, currentView.Id).WhereElementIsNotElementType();
                }
                #endregion

                #region Генерирую фильтры
                // Поиск одинаковых категорий
                if (_entity.What_SameCategory)
                {
                    ElementCategoryFilter sameCatFilter = SelectionSearchFilter.SearchByCategory(_entity.UserSelElem);
                    filters.Add(sameCatFilter);
                }

                // Поиск одинаковых семейств
                if (_entity.What_SameFamily)
                {
                    ElementParameterFilter sameFamFilter = SelectionSearchFilter
                        .SearchByElemBuiltInParam(_entity.UserSelElem, BuiltInParameter.ELEM_FAMILY_PARAM);

                    filters.Add(sameFamFilter);
                }

                // Поиск одинаковых типов
                if (_entity.What_SameType)
                {
                    ElementParameterFilter sameTypeFilter = SelectionSearchFilter
                        .SearchByElemBuiltInParam(_entity.UserSelElem, BuiltInParameter.ELEM_TYPE_PARAM);

                    filters.Add(sameTypeFilter);
                }

                // Поиск по рабочему набору
                if (_entity.What_Workset)
                {
                    ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByElemWorkset(_entity.UserSelElem);
                    filters.Add(sameTypeFilter);
                }

                // Поиск по значению параметра
                if (_entity.What_ParameterData && _entity.What_SelectedParam != null)
                {
                    ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByParamName(doc, _entity.UserSelElem, _entity.What_SelectedParam.RevitParamName);
                    filters.Add(sameTypeFilter);
                }

                #endregion

                LogicalAndFilter combinedFilter = new LogicalAndFilter(filters);
                // Исключаю элементы в группах
                if (_entity.Belong_Group)
                    _resultColl
                        .AddRange(mainFIC.WherePasses(combinedFilter).Where(el => el.GroupId.Equals(ElementId.InvalidElementId))
                        .Select(el => el.Id));
                else
                    _resultColl.AddRange(mainFIC.WherePasses(combinedFilter).ToElementIds());

                // Выделяю элементы
                uiDoc.Selection.SetElementIds(_resultColl);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.Print($"Ошибка попытки выбора подобных. Отправь разработчику: {ex.Message}",
                    MessageType.Error);

                return Result.Cancelled;
            }
        }
    }
}
