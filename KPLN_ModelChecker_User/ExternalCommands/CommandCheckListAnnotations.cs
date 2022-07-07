using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckListAnnotations : IExternalCommand
    {

        /// <summary>
        /// Список категорий для поиска
        /// </summary>
        private List<BuiltInCategory> _bicErrorSearch = new List<BuiltInCategory>
        {
            BuiltInCategory.OST_Lines,
            BuiltInCategory.OST_TextNotes,
            BuiltInCategory.OST_RasterImages,
            BuiltInCategory.OST_GenericAnnotation,
            BuiltInCategory.OST_DetailComponents
        };

        /// <summary>
        /// Список исключений в именах семейств
        /// </summary>
        private List<string> _familyNames = new List<string>
        {
            "022_",
            "023_"
        };

        /// <summary>
        /// Список элементов, которые относятся к ошибкам
        /// </summary>
        private List<ElementId> _errorList = new List<ElementId>();
        
        /// <summary>
        /// Словарь элементов, где ключ - имя листа, значения - аннотации на листе
        /// </summary>
        private Dictionary<ViewSheet, List<ElementId> > _errorDict = new Dictionary<ViewSheet, List<ElementId>>();
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;

            
            
            // Обрабатываю пользовательскую выборку листов
            List<ViewSheet> sheetsList = new List<ViewSheet>();
            List<ElementId> selIds = uidoc.Selection.GetElementIds().ToList(); 
            if (selIds.Count > 0)
            {
                foreach (ElementId selId in selIds)
                {
                    Element elem = doc.GetElement(selId);
                    int catId = elem.Category.Id.IntegerValue;
                    if (catId.Equals((int)BuiltInCategory.OST_Sheets))
                    {
                        ViewSheet curViewSheet = elem as ViewSheet;
                        sheetsList.Add(curViewSheet);
                    }
                }
                if (sheetsList.Count == 0)
                {
                    TaskDialog.Show("Ошибка", "В выборке нет ни одного листа :(", TaskDialogCommonButtons.Ok);
                    return Result.Cancelled;
                }
            }

            // Поиск аннотаций
            try 
            { 
                // Анализирую выбранные листы на количество аннотаций
                if (sheetsList.Count > 0)
                {
                    foreach (ViewSheet viewSheet in sheetsList)
                    {
                        FindAllAnnotationsOnList(doc, viewSheet, _errorDict.GetType());
                    }
                    ShowResult(doc);
                }
                
                // Анализирую все видовые экраны активного листа
                else if (activeView.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Sheets))
                {
                    ViewSheet viewSheet = activeView as ViewSheet;
                    FindAllAnnotationsOnList(doc, viewSheet, _errorList.GetType());
                    ShowResult(uidoc);
                }

                // Анализирую вид
                else
                {
                    FindAllAnnotations(doc, activeView.Id);
                    ShowResult(uidoc);
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для создания фильтра, для игнорирования элементов по имени семейства
        /// </summary>
        private FilteredElementCollector FilteredByStringColl(FilteredElementCollector currentColl)
        {
            foreach (string currentName in _familyNames)
            {
                FilterRule fRule = ParameterFilterRuleFactory.CreateNotBeginsWithRule(new ElementId(BuiltInParameter.ELEM_FAMILY_PARAM), currentName, true);
                ElementParameterFilter eFilter = new ElementParameterFilter(fRule);
                currentColl.WherePasses(eFilter);
            }
            return currentColl;
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на единице выбранного элемента и записи в коллекцию
        /// </summary>
        private void FindAllAnnotations(Document doc, ElementId viewId)
        {
            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(doc, viewId).OfCategory(bic).WhereElementIsNotElementType();
                ICollection<ElementId> collection = FilteredByStringColl(bicColl).ToElementIds();
                _errorList.AddRange(collection);
            }        
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на листе и записи в словарь
        /// </summary>
        private void FindAllAnnotations(Document doc, ElementId viewId, ViewSheet viewSheet)
        {
            foreach (BuiltInCategory bic in _bicErrorSearch)
            {
                FilteredElementCollector bicColl = new FilteredElementCollector(doc, viewId).OfCategory(bic).WhereElementIsNotElementType();
                ICollection<ElementId> collection = FilteredByStringColl(bicColl).ToElementIds();
                if (_errorDict.ContainsKey(viewSheet))
                {
                    _errorDict[viewSheet].AddRange(collection as List<ElementId>);
                }
                else
                {
                    _errorDict.Add(viewSheet, collection as List<ElementId>);
                }
            }
        }

        /// <summary>
        /// Метод для поиска в модели элементов аннотаций на листах и записи в коллекцию или словарь (в зависимости от количества выбранных листов)
        /// </summary>
        private void FindAllAnnotationsOnList(Document doc, ViewSheet viewSheet, Type containerType)
        {
            // Анализирую аннотации на листе
            if (containerType.Equals(typeof(Dictionary<ViewSheet, List<ElementId>>)))
            {
                FindAllAnnotations(doc, viewSheet.Id, viewSheet);
            }
            else
            {
                FindAllAnnotations(doc, viewSheet.Id);
            }

            // Анализирую размещенные виды
            ICollection<ElementId> allViewPorts = viewSheet.GetAllViewports();
            foreach (ElementId vpId in allViewPorts)
            {
                Viewport vp = (Viewport)doc.GetElement(vpId);
                ElementId viewId = vp.ViewId;
                Element currentElement = doc.GetElement(viewId);
                // Анализирую все виды, кроме чертежных видов и легенд
                if (!currentElement.GetType().Equals(typeof(ViewDrafting)) & !currentElement.GetType().Equals(typeof(View)))
                {
                    if (containerType.Equals(typeof(Dictionary<ViewSheet, List<ElementId>>)))
                    { 
                        FindAllAnnotations(doc, viewId, viewSheet);
                    }
                    else
                    {
                        FindAllAnnotations(doc, viewId);
                    }
                }
            }
        }

        /// <summary>
        /// Метод для вывода результатов пользователю, а также для выделения элементов в модели
        /// </summary>
        private void ShowResult(UIDocument uidoc)
        {
            if (_errorList.Count == 0)
            {
                TaskDialog.Show("Результат", "Аннотации не обнаружены :)", TaskDialogCommonButtons.Ok);
            }
            else
            {
                TaskDialog.Show("Результат", $"Аннотации выделены. Количество - {_errorList.Count}", TaskDialogCommonButtons.Ok);
            }
            
            // Выделяю элементы в модели
            uidoc.Selection.SetElementIds(_errorList);
        }

        /// <summary>
        /// Метод для вывода результатов пользователю
        /// </summary>
        private void ShowResult(Document doc)
        {
            if (_errorDict.Count == 0 && _errorList.Count == 0)
            {
                TaskDialog.Show("Результат", "Аннотации не обнаружены :)", TaskDialogCommonButtons.Ok);
            }
            else
            {
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                foreach (KeyValuePair<ViewSheet, List<ElementId> > kvp in _errorDict)
                {
                    foreach (ElementId elementId in kvp.Value)
                    {
                        Element element = doc.GetElement(elementId);
                        WPFDisplayItem item = GetItemByElement(element, element.Name, $"Лист номер {kvp.Key.SheetNumber}", "Данные элементы запрещено использовать на моделируемых видах", Status.Error);
                        item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Id элемента: ", Description = element.Id.ToString() });
                        outputCollection.Add(item);
                    }
                }
                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                List<WPFDisplayItem> sortedOutputCollection = outputCollection.OrderBy(o => o.Header).ToList();
                ObservableCollection<WPFDisplayItem> wpfElements = new ObservableCollection<WPFDisplayItem>();
                int counter = 1;
                foreach (WPFDisplayItem e in sortedOutputCollection)
                {
                    e.Header = string.Format("{0}# {1}", (counter++).ToString(), e.Header);
                    wpfElements.Add(e);
                }
                if (wpfElements.Count != 0)
                {
                    ElementsOutputExtended form = new ElementsOutputExtended(wpfElements, wpfCategories);
                    form.Show();
                }
            }
        }
        
        private WPFDisplayItem GetItemByElement(Element element, string name, string header, string description, Status status)
        {
            StatusExtended exstatus;
            switch (status)
            {
                case Status.Error:
                    exstatus = StatusExtended.Critical;
                    break;
                default:
                    exstatus = StatusExtended.Warning;
                    break;
            }
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus);
            try
            {
                item.SetZoomParams(element, null);
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", element.Category.Name);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Подсказка: ", Description = description });
                HashSet<string> values = new HashSet<string>();
            }
            catch (Exception e)
            {
                try
                {
                    PrintError(e.InnerException);
                }
                catch (Exception) { }
                PrintError(e);
            }
            return item;
        }
    }
}
