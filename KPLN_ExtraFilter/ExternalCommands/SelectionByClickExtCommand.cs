using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Entities;
using KPLN_ExtraFilter.Forms;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SelectionByClickExtCommand : IExternalCommand
    {
        /// <summary>
        /// Итоговая коллекция, которую нужно выделить в модели
        /// </summary>
        private readonly List<ElementId> _resultColl = new List<ElementId>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Пользовательский элемент
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count > 1)
                throw new System.Exception(
                    "Отправь разработчику: Ошибка предварительной проверки. " +
                    "Попало несколько элементов в выборку, хотя должен быть 1.");
            
            // Основные данные для поиска
            Element userSelElem = doc.GetElement(selectedIds.FirstOrDefault());

            // Окно пользовательского ввода
            SelectionByClickForm mainForm = new SelectionByClickForm(doc);
            mainForm.ShowDialog();

            // Поиск элементов комбинированием фильтров
            List<ElementFilter> filters = new List<ElementFilter>();
            if (!mainForm.IsRun) 
                return Result.Succeeded;
            
            SelectionByClickEntity userSelExpEntity = mainForm.CurrentSelectionEntity;

            #region Определяю коллекцию для анализа
            FilteredElementCollector mainFIC = null;
            if (userSelExpEntity.Where_Model)
                mainFIC = new FilteredElementCollector(doc).WhereElementIsNotElementType();
            else if (userSelExpEntity.Where_CurrentView)
            {
                View currentView = doc.ActiveView 
                    ?? throw new System.Exception("Отправь разработчику: Не удалось определить открытый вид");

                mainFIC = new FilteredElementCollector(doc, currentView.Id).WhereElementIsNotElementType();
            }
            #endregion

            #region Генерирую фильтры
            // Поиск одинаковых категорий
            if (userSelExpEntity.What_SameCategory)
            {
                ElementCategoryFilter sameCatFilter = SelectionSearchFilter.SearchByCategory(userSelElem);
                filters.Add(sameCatFilter);
            }

            // Поиск одинаковых семейств
            if (userSelExpEntity.What_SameFamily)
            {
                ElementParameterFilter sameFamFilter = SelectionSearchFilter
                    .SearchByElemParam(userSelElem, BuiltInParameter.ELEM_FAMILY_PARAM);
                
                filters.Add(sameFamFilter);
            }

            // Поиск одинаковых типов
            if (userSelExpEntity.What_SameType)
            {
                ElementParameterFilter sameTypeFilter = SelectionSearchFilter
                    .SearchByElemParam(userSelElem, BuiltInParameter.ELEM_TYPE_PARAM);
                
                filters.Add(sameTypeFilter);
            }

            // Поиск по рабочему набору
            if (userSelExpEntity.What_Workset)
            {
                ElementFilter sameTypeFilter = SelectionSearchFilter.SearchByWorkset(userSelElem);
                filters.Add(sameTypeFilter);
            }

            #endregion

            LogicalAndFilter combinedFilter = new LogicalAndFilter(filters);
            // Исключаю элементы в группах
            if (userSelExpEntity.Belong_Group)
                _resultColl
                    .AddRange(mainFIC.WherePasses(combinedFilter).Where(el => el.GroupId.IntegerValue == -1)
                    .Select(el => el.Id));
            else
                _resultColl.AddRange(mainFIC.WherePasses(combinedFilter).ToElementIds());

            // Выделяю элементы
            uidoc.Selection.SetElementIds(_resultColl);

            return Result.Succeeded;
        }
    }
}
