using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckFamilies : IExternalCommand
    {
        private delegate void FormRunner(ProgressBar progressBar);

        private void ProgressBarRun(ProgressBar progressBar)
        {
            progressBar.Show();
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Application app = commandData.Application.Application;
            app.FailuresProcessing += FailuresProcessor;

            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                List<Element> famColl = new FilteredElementCollector(doc).OfClass(typeof(Family)).ToList();
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();

                //ProgressBar progressBar = new ProgressBar(
                //    famColl.Count,
                //    "Проверка семейств",
                //    "Проверка семейств и типоразмеров на дубликаты, а также проверка расположения семейств",
                //    true,
                //    "Подсказка",
                //    "Дубликаты семейств приводят к утяжелению файла, а также к сложностям в администрировании и редактировании проекта в целом." +
                //    "Семейства с запрещенных источников приводят к сложностям в администрировании и создании проекта в целом");
                
                //progressBar.Show();
                foreach (Family currentFam in famColl)
                {
                    //progressBar.CurrentProgress++;

                    CheckFamilyName(currentFam, famColl, ref outputCollection);
                    CheckFamilyPath(doc, currentFam, ref outputCollection);
                }
                //progressBar.Close();

                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });

                IEnumerable<WPFDisplayItem> distCategories = outputCollection.GroupBy(w => w.CategoryId).Select(g => g.First());
                foreach (WPFDisplayItem item in distCategories)
                {
                    int count = outputCollection.Where(x => x.Category.Equals(item.Category)).Count();

                    Family entity = outputCollection.FirstOrDefault(x => x.Equals(item)).Element as Family;
                    if (entity == null)
                    {
                        FamilySymbol fSymbol = outputCollection.FirstOrDefault(x => x.Equals(item)).Element as FamilySymbol;
                        entity = fSymbol.Family;
                    }
                    Category category = entity.FamilyCategory;

                    wpfCategories.Add(new WPFDisplayItem(category.Id.IntegerValue, StatusExtended.Critical)
                    {
                        Name = $"{category.Name} ({count})"
                    });
                }

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
                else
                {
                    Print("[Наименование] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    if (e.InnerException.InnerException != null)
                        PrintError(e.InnerException.InnerException);
                    else
                        PrintError(e.InnerException);
                else
                    PrintError(e);

                return Result.Failed;
            }
            finally
            {
                app.FailuresProcessing -= FailuresProcessor;
            }
        }

        private void FailuresProcessor(object sender, Autodesk.Revit.DB.Events.FailuresProcessingEventArgs e)
        {
            FailuresAccessor fAcc = e.GetFailuresAccessor();
            List<FailureMessageAccessor> failureMessageAccessors = fAcc.GetFailureMessages().ToList();
            if (failureMessageAccessors.Count > 0)
            {
                List<ElementId> elemsToDelete = new List<ElementId>();
                foreach (FailureMessageAccessor fma in failureMessageAccessors)
                {
                    Document fDoc = fAcc.GetDocument();
                    //elemsToDelete.AddRange(fma.GetFailingElementIds());

                    List<ElementId> fmFailElemsId = fma.GetFailingElementIds().ToList();
                    foreach (ElementId elId in fmFailElemsId)
                    {
                        Element fmFailElem = fDoc.GetElement(elId);
                        Type fmType = fmFailElem.GetType();
                        if (!fmType.Equals(typeof(PlanarFace))
                            && !fmType.Equals(typeof(ReferencePlane)))
                        {
                            elemsToDelete.Add(elId);
                        }
                    }
                }

                fAcc.DeleteAllWarnings();
                if (elemsToDelete.Count > 0)
                {
                    try
                    {
                        fAcc.DeleteElements(elemsToDelete);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException) 
                    {
                        e.SetProcessingResult(FailureProcessingResult.Continue);
                        return;
                    }

                    e.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
                    return;
                }

                e.SetProcessingResult(FailureProcessingResult.Continue);
            }
        }

        /// <summary>
        /// Проверка имен семейства и его типоразмеров
        /// </summary>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="docFamilies">Коллекция семейств проекта</param>
        /// <param name="outputCollection">Коллекция элементов WPFDisplayItem для отчета</param>
        private void CheckFamilyName(Family currentFam, List<Element> docFamilies, ref ObservableCollection<WPFDisplayItem> outputCollection)
        {
            List<Element> currentFamilySymols = new List<Element>();
            string currentFamName = currentFam.Name;
            if (Regex.Match(currentFamName, @"\b[.0]\d*$").Value.Length > 2)
            {
                WPFDisplayItem item = GetItemByElement(
                    currentFam,
                    $"{currentFamName}",
                    "Предупреждение семейства",
                    $"Данное семейство - это резервная копия. Запрещено использовать резервные копии!",
                    Status.Error,
                    null);

                item.Collection.Add(
                    new WPFDisplayItem(-1, StatusExtended.Critical)
                    {
                        Header = "Инфо:",
                        Description = "Необходимо корректно обновить семейство. Резервные копии - могут содержать не корректную информацию!"
                    });

                outputCollection.Add(item);
            }

            string similarFamilyName = SearchSimilarName(currentFamName, docFamilies);
            if (!similarFamilyName.Equals(String.Empty))
            {
                WPFDisplayItem item = GetItemByElement(
                    currentFam,
                    $"{currentFamName}",
                    "Предупреждение семейства",
                    $"Возможно семейство является копией семейства «{similarFamilyName}»",
                    Status.Error,
                    null);

                item.Collection.Add(
                    new WPFDisplayItem(-1, StatusExtended.Critical)
                    {
                        Header = "Инфо:",
                        Description = "Копий семейств в проекте быть не должно!"
                    });

                outputCollection.Add(item);
            }

            foreach (ElementId id in currentFam.GetFamilySymbolIds())
            {
                FamilySymbol symbol = currentFam.Document.GetElement(id) as FamilySymbol;
                currentFamilySymols.Add(symbol);
            }

            foreach (FamilySymbol currentSymbol in currentFamilySymols)
            {
                string currentSymName = currentSymbol.Name;
                string similarSymbolName = SearchSimilarName(currentSymName, currentFamilySymols);

                if (!similarSymbolName.Equals(String.Empty))
                {
                    WPFDisplayItem item = GetItemByElement(
                        currentSymbol,
                        $"{currentFamName}: {currentSymName}",
                        "Предупреждение типоразмера",
                        $"Возможно тип является копией типоразмера «{similarSymbolName}»",
                        Status.Error,
                        null);

                    item.Collection.Add(
                        new WPFDisplayItem(-1, StatusExtended.Critical)
                        {
                            Header = "Инфо:",
                            Description = "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!"
                        });

                    outputCollection.Add(item);
                }
            }
        }

        /// <summary>
        /// Проверка пути к семейству
        /// </summary>
        /// <param name="doc">Файл Revit</param>
        /// <param name="currentFam">Семейство для проверки</param>
        /// <param name="outputCollection">Коллекция элементов WPFDisplayItem для отчета</param>
        private void CheckFamilyPath(Document doc, Family currentFam, ref ObservableCollection<WPFDisplayItem> outputCollection)
        {
            BuiltInCategory currentBIC;
            Category currentCat = currentFam.FamilyCategory;
            if (currentCat != null)
                currentBIC = (BuiltInCategory)currentCat.Id.IntegerValue;
            else
                return;

            if (currentFam.get_Parameter(BuiltInParameter.FAMILY_SHARED).AsInteger() != 1
                && currentFam.IsEditable
                && !currentCat.CategoryType.Equals(CategoryType.Annotation)
                && !currentBIC.Equals(BuiltInCategory.OST_ProfileFamilies)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponents)
                && !currentBIC.Equals(BuiltInCategory.OST_GenericAnnotation)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponentsHiddenLines)
                && !currentBIC.Equals(BuiltInCategory.OST_DetailComponentTags))
            {
                Document famDoc;
                try
                {
                    famDoc = doc.EditFamily(currentFam);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Работа остановлена, т.к. семейство {currentFam.Name} не может быть открыто. Причина: {ex}");
                }
                
                string famPath = famDoc.PathName;
                if (!famPath.StartsWith("X:\\")
                    && !famPath.Contains("03_Скрипты")
                    && !famPath.ToLower().Contains("ostec")
                    && !famPath.Contains("KPLN_Loader")
                    && famDoc.IsFamilyDocument == true)
                {
                    WPFDisplayItem item = GetItemByElement(
                        currentFam,
                        $"{currentFam.Name}",
                        "Предупреждение источника семейства",
                        $"Данное семейство - не с диска Х. Запрещено использовать сторонние источники!",
                        Status.Error,
                        null);

                    string descr = string.Empty;
                    if (!string.IsNullOrEmpty(famPath)
                        && !famPath.Contains("KPLN_Loader"))
                        descr = $"Текущий путь к семейству: {famPath}. Использовать в проекте данное семейство можно только по согласованию в BIM-отделе.";
                    else
                        descr = "Источник на сервере - не определен. Использовать в проекте данное семейство можно только по согласованию в BIM-отделе.";

                    item.Collection.Add(
                        new WPFDisplayItem(-1, StatusExtended.Critical)
                        {
                            Header = "Инфо:",
                            Description = descr
                        });

                    outputCollection.Add(item);
                }

                famDoc.Close(false);
            }
        }

        private WPFDisplayItem GetItemByElement(Element element, string name, string header, string description, Status status, BoundingBoxXYZ box)
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

            int catId;
            string catName = string.Empty;
            Category cat = element.Category;
            if (cat != null)
            {
                catId = element.Category.Id.IntegerValue;
                catName = element.Category.Name;
            }
            else
            {
                Family family = (Family)element;
                catId = family.FamilyCategory.Id.IntegerValue;
                catName = family.FamilyCategory.Name;
            }

            WPFDisplayItem item = new WPFDisplayItem(catId, exstatus, "✔");
            try
            {
                item.Element = element;
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", catName);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(catId, exstatus) { Header = "Подсказка: ", Description = description });
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

        private bool CatInList(List<List<object>> aCats, Category cat)
        {
            foreach (List<object> c in aCats)
            {
                Category ca = c[1] as Category;
                if (ca.Id.IntegerValue == cat.Id.IntegerValue)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Поиск похожего имени. Одинаковым должна быть только первичная часть имени, до среза по циферным значениям
        /// </summary>
        /// <param name="currentName">Имя, которое нужно проанализировать</param>
        /// <param name="elemsColl">Коллекция, по которой нужно осуществлять поиск</param>
        /// <returns>Имя подобного элемента</returns>
        private string SearchSimilarName(string currentName, List<Element> elemsColl)
        {
            string similarFamilyName = String.Empty;

            // Осуществляю поиск цифр в конце имени
            string digitEndTrimmer = Regex.Match(currentName, @"\d*$").Value;
            // Осуществляю срез имени на найденные цифры в конце имени
            string truePartOfName = currentName.TrimEnd(digitEndTrimmer.ToArray());
            if (digitEndTrimmer.Length > 0)
            {
                foreach (Element checkElem in elemsColl)
                {
                    if (!checkElem.Equals(currentName) && checkElem.Name.Equals(truePartOfName.TrimEnd(new char[] { ' ' })))
                    {
                        similarFamilyName = checkElem.Name;
                    }
                }
            }
            return similarFamilyName;
        }



    }
}
