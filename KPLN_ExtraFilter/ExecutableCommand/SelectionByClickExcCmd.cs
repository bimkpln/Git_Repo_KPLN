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
        private readonly SelectionByClickM _soEntity;

        /// <summary>
        /// Итоговая коллекция, которую нужно выделить в модели
        /// </summary>
        private readonly List<ElementId> _resultColl = new List<ElementId>();

        public SelectionByClickExcCmd(SelectionByClickM soEntity)
        {
            _soEntity = soEntity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            Document doc = uiDoc.Document;

            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(SelectionByClickExtCmd.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
            try
            {
                // Поиск элементов комбинированием фильтров
                List<ElementFilter> filters = new List<ElementFilter>();

                #region Определяю коллекцию для анализа
                FilteredElementCollector mainFIC = null;
                if (_soEntity.Where_Model)
                    mainFIC = new FilteredElementCollector(doc).WhereElementIsNotElementType();
                else if (_soEntity.Where_CurrentView)
                {
                    View currentView = doc.ActiveView
                        ?? throw new System.Exception("Отправь разработчику: Не удалось определить открытый вид");

                    mainFIC = new FilteredElementCollector(doc, currentView.Id).WhereElementIsNotElementType();
                }
                #endregion

                #region Генерирую фильтры
                // Поиск одинаковых категорий
                if (_soEntity.What_SameCategory)
                {
                    ElementCategoryFilter sameCatFilter = SelectionSearchFilter.SearchByCategory(_soEntity.UserSelElem);
                    filters.Add(sameCatFilter);
                }

                // Поиск одинаковых семейств
                if (_soEntity.What_SameFamily)
                {
                    ElementParameterFilter sameFamFilter = SelectionSearchFilter
                        .SearchByElemBuiltInParam(_soEntity.UserSelElem, BuiltInParameter.ELEM_FAMILY_PARAM);

                    filters.Add(sameFamFilter);
                }

                // Поиск одинаковых типов
                if (_soEntity.What_SameType)
                {
                    ElementParameterFilter sameTypeFilter = SelectionSearchFilter
                        .SearchByElemBuiltInParam(_soEntity.UserSelElem, BuiltInParameter.ELEM_TYPE_PARAM);

                    filters.Add(sameTypeFilter);
                }

                // Поиск по рабочему набору
                if (_soEntity.What_Workset)
                {
                    ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByWorkset(_soEntity.UserSelElem);
                    filters.Add(sameTypeFilter);
                }

                // Поиск по значению параметра
                if (_soEntity.What_ParameterData && _soEntity.What_SelectedParam != null)
                {
                    ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByParamName(doc, _soEntity.UserSelElem, _soEntity.What_SelectedParam.CurrentParamName);
                    filters.Add(sameTypeFilter);
                }

                #endregion

                LogicalAndFilter combinedFilter = new LogicalAndFilter(filters);
                // Исключаю элементы в группах
                if (_soEntity.Belong_Group)
                    _resultColl
                        .AddRange(mainFIC.WherePasses(combinedFilter).Where(el => el.GroupId.IntegerValue == -1)
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
