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
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus, "✔");
            try
            {
                item.SetZoomParams(element, box);
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
        private void IncreaseCategoryAmmount(List<List<object>> aCats, Category cat)
        {
            foreach (List<object> c in aCats)
            {
                Category ca = c[1] as Category;
                if (ca.Id.IntegerValue == cat.Id.IntegerValue)
                {
                    c[0] = (int)c[0] + 1;
                }
            }
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                HashSet<int> _fam_ids = new HashSet<int>();
                HashSet<string> _fam_names = new HashSet<string>();
                List<Family> _fams = new List<Family>();
                List<List<object>> aCats = new List<List<object>>();
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                foreach (Family family in new FilteredElementCollector(doc).OfClass(typeof(Family)).ToElements())
                {
                    if (!CatInList(aCats, family.FamilyCategory))
                    {
                        aCats.Add(new List<object> { 0, family.FamilyCategory });
                    }
                    if (!_fam_ids.Contains(family.Id.IntegerValue))
                    {
                        _fam_ids.Add(family.Id.IntegerValue);
                        _fams.Add(family);
                        _fam_names.Add(family.Name);
                    }
                    else
                    {
                        continue;
                    }
                }
                foreach (Family family in _fams)
                {
                    try
                    {
                        HashSet<string> _sym_names = new HashSet<string>();
                        List<FamilySymbol> _syms = new List<FamilySymbol>();
                        string _family_name = family.Name;
                        if (IsInteger(_family_name[_family_name.Length - 1]) && !IsInteger(_family_name[_family_name.Length - 2]))
                        {
                            foreach (int i in new int[] { 1, 2 })
                            {
                                if (_fam_names.Contains(GetShortenedName(_family_name, i)))
                                {
                                    try
                                    {
                                        IncreaseCategoryAmmount(aCats, family.Category);
                                        WPFDisplayItem item = GetItemByElement(family, family.Name, "Предупреждение семейства", string.Format("Возможно семейство является копией семейства «{0}.rfa»", GetShortenedName(_family_name, i)), Status.Error, null);
                                        item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Инфо:", Description = "Копий семейств в проекте быть не должно! Копии типоразмеров - допускаются, но только по согласованию с BIM-отделом" });
                                        outputCollection.Add(item);
                                    }
                                    catch (Exception)
                                    { }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        foreach (ElementId id in family.GetFamilySymbolIds())
                        {
                            FamilySymbol symbol = family.Document.GetElement(id) as FamilySymbol;
                            _syms.Add(symbol);
                            _sym_names.Add(symbol.Name);
                        }
                        HashSet<string> _check_names = new HashSet<string>();
                        foreach (FamilySymbol symbol in _syms)
                        {
                            string _symbol_name = symbol.Name;
                            if (IsInteger(_symbol_name[_symbol_name.Length - 1]) && !IsInteger(_symbol_name[_symbol_name.Length - 2]))
                            {
                                foreach (int i in new int[] { 1, 2 })
                                {
                                    _check_names.Add(GetShortenedName(_symbol_name, i));
                                    if (_sym_names.Contains(GetShortenedName(_symbol_name, i)))
                                    {
                                        try
                                        {
                                            IncreaseCategoryAmmount(aCats, symbol.Category);
                                            WPFDisplayItem item = GetItemByElement(symbol, string.Format("{0} ({1})", _family_name, _symbol_name), "Предупреждение типоразмера", string.Format("Возможно тип является копией типоразмера «{0}»", GetShortenedName(_symbol_name, i)), Status.Error, null);
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Инфо:", Description = "Копии необходимо наименовывать корректно, либо избегать появления копий в проекте!" });
                                            outputCollection.Add(item);
                                        }
                                        catch (Exception)
                                        { }
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
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
        private static string GetShortenedName(string value, int ammount = 1)
        {
            string result = string.Empty;
            for (int i = 0; i < value.Length - ammount; i++)
            {
                try
                {
                    result += value[i];
                }
                catch (Exception)
                { }
            }
            return result;
        }
        private static bool IsInteger(char c)
        {
            if("0123456789".Contains(c))
            { 
                return true; 
            }
            return false;

        }
    }
}
