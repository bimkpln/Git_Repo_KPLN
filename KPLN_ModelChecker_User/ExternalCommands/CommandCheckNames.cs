using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckNames : IExternalCommand
    {
        
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
            try 
            {
                catId = element.Category.Id.IntegerValue;
                catName = element.Category.Name;
            }
            catch
            {
                Family family = (Family)element;
                catId = family.FamilyCategory.Id.IntegerValue;
                catName = family.FamilyCategory.Name;
            }
            WPFDisplayItem item = new WPFDisplayItem(catId, exstatus, "✔");
            try
            {
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
            foreach(List<object> c in aCats)
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
                    if (!checkElem.Equals(currentName) && checkElem.Name.Equals(truePartOfName.TrimEnd(new char[] { ' ', '.' })))
                    {
                        similarFamilyName = checkElem.Name;
                    }
                }
            }
            return similarFamilyName;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                List<Element> docFamilies = new List<Element>();
                HashSet<int> familyIds = new HashSet<int>();
                HashSet<string> docFamilyNames = new HashSet<string>();
                List<List<object>> aCats = new List<List<object>>();
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                
                foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).ToElements())
                {
                    if (!CatInList(aCats, family.FamilyCategory))
                    {
                        aCats.Add(new List<object> { 0, family.FamilyCategory });
                    }
                    if (!familyIds.Contains(family.Id.IntegerValue))
                    {
                        docFamilies.Add(family);
                        familyIds.Add(family.Id.IntegerValue);
                        docFamilyNames.Add(family.Name);
                    }
                    else
                    {
                        continue;
                    }
                }
                
                foreach (Family currentFam in docFamilies)
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
                            null
                        );

                        item.Collection.Add(
                            new WPFDisplayItem(-1, StatusExtended.Critical)
                            {
                                Header = "Инфо:",
                                Description = "Необходимо корректно обновить семейство. Резервные копии - могут содержать не корректную информацию!"
                            }
                        );

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
                            null
                        );
                                
                        item.Collection.Add(
                            new WPFDisplayItem(-1, StatusExtended.Critical) { 
                                Header = "Инфо:",
                                Description = "Копий семейств в проекте быть не должно!" 
                            }
                        );
                                
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
                                null
                            );
                                    
                            item.Collection.Add(
                                new WPFDisplayItem(-1, StatusExtended.Critical) { 
                                    Header = "Инфо:",
                                    Description = "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!"
                                }
                            );
                                    
                            outputCollection.Add(item);
                        }
                    }
                }
                
                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                foreach (List<object> cat in aCats)
                {
                    if ((int)cat[0] == 0) { continue; }
                    Category category = cat[1] as Category;
                    wpfCategories.Add(new WPFDisplayItem(category.Id.IntegerValue, StatusExtended.Critical) { Name = string.Format("{0} ({1})", category.Name, ((int)cat[0]).ToString()) });
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
                try
                {
                    PrintError(e.InnerException);
                    try
                    {
                        PrintError(e.InnerException.InnerException);
                    }
                    catch (Exception)
                    { }
                }
                catch (Exception)
                { }
                PrintError(e);
                return Result.Failed;
            }
        }

    }
}
