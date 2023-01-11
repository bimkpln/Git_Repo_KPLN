using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;
using KPLN_Quantificator.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Quantificator
{
    public class QFolder
    { 
        public string Name { get; set; }
        public int Parent { get; set; }
        public int Index { get; set; }

        public List<QFolder> Folders = new List<QFolder>();

        public List<ItemsSearch> SearchList = new List<ItemsSearch>();

        public void FindParentFolder(List<QFolder> folder_list)
        {
            foreach (QFolder folder in folder_list)
            { 
                if (folder.Name == APITools.GetSingleSplittedValue(Name, '.'))
                {
                    folder.Folders.Add(this);
                    //Parent = folder.Index;
                    return;
                }        
            }

        }
    }
    public class ItemsSearch
    { 
        public string Name { get; set; }
        public int Index { get; set; }

        public List<Guid> Guids = new List<Guid>();


    }
    class APITools
    {
        public static ItemsSearch CreateNewSearch(SavedItem item)
        {
            ModelItemCollection model_collection = ((SelectionSet)item).GetSelectedItems();
            ItemsSearch item_search = new ItemsSearch();
            item_search.Name = item.DisplayName;
            item_search.Index = GlobalPreferences.GetItemIndex();
            if (model_collection != null)
            {
                foreach (ModelItem model_item in model_collection)
                {
                    Guid item_guid = FindGuid(model_item);
                    if (item_guid.ToString() != "00000000-0000-0000-0000-000000000000")
                    {
                        item_search.Guids.Add(item_guid);
                    }
                }
            }
            return item_search;
        }
        private static Guid FindGuid(ModelItem mditem)
        {
            Guid item_guid = Guid.Parse("00000000-0000-0000-0000-000000000000");
            if (mditem.Parent != null)
            {
                if (mditem.InstanceGuid.ToString() == "00000000-0000-0000-0000-000000000000")
                {
                    item_guid = FindGuid(mditem.Parent);
                }
                else
                {
                    item_guid = mditem.InstanceGuid;
                }
            }
            return item_guid;
        }
        public static List<SavedItem> GetSubItems(SavedItem item)//Вызов из GetSavedItems()
        {
            List<SavedItem> all_saved_items = new List<SavedItem>();
            foreach (SavedItem child_item in ((GroupItem)item).Children)
            {
                if (child_item.IsGroup)
                {
                    foreach (SavedItem i in GetSubItems(child_item))
                    {
                        all_saved_items.Add(i);
                    }
                }
                else 
                { 
                    all_saved_items.Add(child_item);
                }
            }
            return all_saved_items;
        }
        private static List<SavedItem> GetSubFolders (SavedItem item)//Вызов из GetSavedItems()
        {
            List<SavedItem> all_saved_items = new List<SavedItem>();
            foreach (SavedItem child_item in ((GroupItem)item).Children)
            {
                if (child_item.IsGroup)
                {
                    all_saved_items.Add(child_item);
                    foreach (SavedItem i in GetSubFolders(child_item))
                    {
                        all_saved_items.Add(i);
                    }
                }
            }
            return all_saved_items;
        }
        public static List<string> GetAllSaveditemsNames()
        {
            List<string> items = new List<string>();
            //
            List<SavedItem> all_saved_items = new List<SavedItem>();
            foreach (SavedItem item in Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup)
                {
                    all_saved_items.Add(item);
                    foreach (SavedItem child_item in GetSubFolders(item))
                    {
                        all_saved_items.Add(child_item);
                    }
                    //Print(string.Format("GROUP: {0}", item.DisplayName)); 
                }
            }
            foreach (SavedItem item in all_saved_items)
            {
                items.Add(item.DisplayName);
            }
            items.Sort();
            //
            return items;
        }
        public static List<SavedItem> GetSavedItemGroups()
        {
            List<SavedItem> all_saved_groups = new List<SavedItem>();
            foreach (SavedItem item in Autodesk.Navisworks.Api.Application.ActiveDocument.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup)
                {
                    all_saved_groups.Add(item);
                    foreach (SavedItem child_item in GetSubFolders(item))
                    {
                        all_saved_groups.Add(child_item);
                    }
                }
            }

            return all_saved_groups;
        }
        public static List<SavedItem> GetSavedItems(string folder_name)
        {
            List<SavedItem> all_saved_items = new List<SavedItem>();
            try
            {
                foreach (SavedItem item in GetSavedItemGroups())
                {
                    if (item.DisplayName == folder_name)
                    {
                         if (item.IsGroup) 
                         {
                            foreach (SavedItem child_item in GetSubItems(item))
                            {
                                all_saved_items.Add(child_item);
                            }
                         }
                         else 
                         {
                             all_saved_items.Add(item);
                         }                   
                    }

                }
            }
            catch (Exception) { }
            return all_saved_items;
        }
        private static DocumentSelectionSets GetDocumentSelectionSets()
        {
            Document active_document = Autodesk.Navisworks.Api.Application.ActiveDocument;
            DocumentSelectionSets selection_sets = active_document.SelectionSets;
            return selection_sets;
        }
        public static FolderItem CreateNewFolder(string folder_name)
        {
            try
            {
                DocumentSelectionSets selection_sets = GetDocumentSelectionSets();
                Int32 folder_index = selection_sets.Value.IndexOfDisplayName(folder_name);
                if (folder_index == -1)
                {
                    selection_sets.AddCopy(new FolderItem() { DisplayName = folder_name });
                }
                FolderItem folder_item = selection_sets.Value[selection_sets.Value.IndexOfDisplayName(folder_name)] as FolderItem;
                return folder_item;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static List<string> GetUniqValuesOfDataProperty(List<ModelItem> elements, string category, string property)
        {
            List<string> value_list = new List<string>();

            foreach (ModelItem item in elements)
            {
                foreach (PropertyCategory prop_cat in item.PropertyCategories)
                {
                    if (prop_cat.DisplayName == category)
                    {
                        foreach (DataProperty data_prop in prop_cat.Properties)
                        {
                            if (data_prop.DisplayName == property)
                            {
                                try
                                {
                                    if (!value_list.Contains(data_prop.Value.ToDisplayString()))
                                    {
                                        value_list.Add(data_prop.Value.ToDisplayString());
                                    }
                                }
                                catch (Exception)
                                {
                                    Output.PrintAlert(string.Format("Элемент «{0}» : отсутствует один или несколько параметров;", item.DisplayName));
                                }


                            }
                        }
                    }

                }
            }
            return value_list;
        }
        public static HashSet<string> GetSplittedValues(List<string> values, char split_by)
        {
            HashSet<string> end_values = new HashSet<string>();
            foreach (string value in values)
            {
                if (value.Contains(split_by))
                {
                    HashSet<string> split_values = GetSplittedValue(value, split_by);
                    if (split_values.Count != 0)
                    {
                        foreach (string split_value in split_values)
                        {
                            end_values.Add(split_value);
                        }
                    }
                }
            }
            return end_values;
        }
        public static HashSet<string> GetSplittedValue(string input_string, char split_by)
        {
            HashSet<string> result = new HashSet<string>();
            string[] parts = input_string.Split(split_by);
            List<string> split_value_parts = new List<string>();
            int step = 0;
            foreach (string part in parts)
            {
                step++;
                if (step != parts.Length)
                {
                    split_value_parts.Add(part);
                }
            }
            string value_string = string.Join(".", split_value_parts);
            result.Add(value_string);
            if (value_string.Contains(split_by))
            {
                foreach (string sub_value in GetSplittedValue(value_string, split_by))
                {
                    result.Add(sub_value);
                }
            }
            return result;
        }       
        public static string GetSingleSplittedValue(string input_string, char split_by)
        {
            string[] parts = input_string.Split(split_by);
            List<string> split_value_parts = new List<string>();
            int step = 0;
            foreach (string part in parts)
            {
                step++;
                if (step != parts.Length)
                {
                    split_value_parts.Add(part);
                }
            }
            string value_string = string.Join(".", split_value_parts);
            return value_string;
        }
        public static string GetParameterValue(SavedItem item, string category, string parameter)
        {
            try
            {
                ModelItemCollection model_collection = ((SelectionSet)item).GetSelectedItems();
                foreach (ModelItem model_item in model_collection)
                {
                    foreach (PropertyCategory cat in model_item.PropertyCategories)
                    {
                        if (cat.DisplayName == category)
                        {
                            foreach (DataProperty prop in cat.Properties)
                            {
                                if (prop.DisplayName == parameter)
                                {
                                    try
                                    {
                                        return prop.Value.ToDisplayString();
                                    }
                                    catch (Exception)
                                    {
                                        continue;
                                    }
                                   
                                }
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception)
            {
                return null;
            }

        }
        public static bool ItemByWBSInList(List<DBItemGroup> items, string wbs)
        {
            foreach (DBItemGroup item in items)
            {
                if (item.GlobalWBS == wbs)
                {
                    return true;
                }
            }
            return false;
        }
        public static SelectionSet GetSelectionSetByProperty(string category, string property, string value, string name)
        {
            try
            {
                DocumentSelectionSets selection_sets = GetDocumentSelectionSets();
                Search search = new Search();
                search.Locations = SearchLocations.DescendantsAndSelf;
                search.Selection.SelectAll();
                SearchCondition search_condition = SearchCondition.HasPropertyByDisplayName(category, property);
                search.SearchConditions.Add(search_condition.EqualValue(VariantData.FromDisplayString(value)));
                SelectionSet selection_set = new SelectionSet(search) { DisplayName = name };
                selection_sets.AddCopy(selection_set);
                return selection_set;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static bool FindSelectionSetParent(SelectionSet item, List<FolderItem> folders)
        {
            foreach (FolderItem folder in folders)
            {
                if (folder.DisplayName != item.DisplayName)
                {
                    if (folder.DisplayName == GetSingleSplittedValue(item.DisplayName, '.'))
                    {
                        MoveSelectionSetToFolder(item, folder);
                        return true;
                    }
                }
            }
            foreach (FolderItem folder in folders)
            {
                if (folder.DisplayName != item.DisplayName)
                {
                    if (item.DisplayName.Contains(folder.DisplayName))
                    {
                        MoveSelectionSetToFolder(item, folder);
                        return true;
                    }
                }
            }
            return false;
        }
        public static void MoveSelectionSetToFolder(SelectionSet selection_set, FolderItem folder)
        {
            DocumentSelectionSets selection_sets = GetDocumentSelectionSets();
            SavedItem saved_item = selection_sets.Value[selection_sets.Value.IndexOfDisplayName(selection_set.DisplayName)] as SavedItem;
            selection_sets.Move(saved_item.Parent, selection_sets.Value.IndexOfDisplayName(selection_set.DisplayName), folder, 0);
        }
        public static bool MoveFolderToParent(FolderItem item, List<FolderItem> folders)
        {
            foreach (FolderItem folder in folders)
            {
                if (folder.DisplayName != item.DisplayName)
                {
                    if (folder.DisplayName == GetSingleSplittedValue(item.DisplayName, '.'))
                    {
                        MoveFolderToFolder(item, folder);
                        return true;
                    }
                }
            }
            foreach (FolderItem folder in folders)
            {
                if (folder.DisplayName != item.DisplayName)
                {
                    if (item.DisplayName.Contains(folder.DisplayName))
                    {
                        MoveFolderToFolder(item, folder);
                        return true;
                    }
                }
            }
            return false;
        }
        public static void MoveFolderToFolder(FolderItem child_folder, FolderItem parent_folder)
        {
            DocumentSelectionSets selection_sets = GetDocumentSelectionSets();
            FolderItem child_folder_item = selection_sets.Value[selection_sets.Value.IndexOfDisplayName(child_folder.DisplayName)] as FolderItem;
            selection_sets.Move(child_folder_item.Parent, selection_sets.Value.IndexOfDisplayName(child_folder.DisplayName), parent_folder, 0);
        }
    }
}
