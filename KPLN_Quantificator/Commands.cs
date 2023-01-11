using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Data;
using Autodesk.Navisworks.Api.Takeoff;
using System.Linq;
using System.Threading;
using KPLN_Quantificator.Forms;

namespace KPLN_Quantificator
{
    public class Commands
    {
        public static void GroupClashes() { }

        public static void AddQuantItems(string folder_name, string name_cat, string name_prop, string description_cat, string description_prop)
        {
            Output.PrintHeader("KPLN Extention : Создание элементов для каталога элементов (Quantification)");
            List<SavedItem> saved_items = APITools.GetSavedItems(folder_name);
            GlobalPreferences.Update();
            int lastWbs = 0;
            foreach (SavedItem saved_item in saved_items)
            {
                SelectionSet selectionSet = saved_item as SelectionSet;
                if (selectionSet.GetSelectedItems().Count == 0)
                {
                    Output.PrintAlert($"Поисковый набор «{saved_item.DisplayName}» не содержит элементов в модели");
                    continue;
                }
                
                // Поиск элементов вне списка - нужно переделать, т.к. раньше привязка была на имя поиск. набора
                if (!APITools.ItemByWBSInList(GlobalPreferences.items, saved_item.DisplayName))
                {
                    string name = APITools.GetParameterValue(saved_item, name_cat, name_prop);
                    string description = APITools.GetParameterValue(saved_item, description_cat, description_prop);

                    if (description != null && name != null)
                    {
                        DBItemGroup parent_folder = DBTools.GetParentByNameContains(saved_item.DisplayName);
                        if (parent_folder != null)
                        {
                            int parent_wbs = DBTools.GetLastWbs(parent_folder.Id);
                            if (lastWbs == 0)
                                lastWbs = parent_wbs++;
                            DBTools.GPInsertItem(parent_folder, saved_item, (++lastWbs).ToString());
                            long lastInsert = DBTools.InsertItem(parent_folder.Id, name, description, (lastWbs).ToString());

                            List<DBObject> objects = DBTools.GetObjectsByParentId(parent_folder.Id);
                            ItemsSearch item_search = APITools.CreateNewSearch(saved_item);
                            foreach (Guid search_guid in item_search.Guids)
                            {
                                if (!DBTools.ItemInList(objects, search_guid))
                                {
                                    string wbs = DBTools.GetNextWbs(objects);
                                    DBTools.DoTakeoff(lastInsert, search_guid, lastWbs.ToString());
                                    DBTools.GPInsertObject(parent_folder, saved_item, objects, lastWbs.ToString());
                                }
                            }

                            Output.PrintSuccess(string.Format("Элемент «{0}» добавлен в Quantification;", saved_item.DisplayName));
                        }
                        else
                            Output.PrintAlert(string.Format("Элемент «{0}» : не найдена родительская папка;", saved_item.DisplayName));
                    }
                    else
                    {
                        Output.PrintAlert($"Элемент «{saved_item.DisplayName}»: " +
                            $"отсутствует/не заполнен один или несколько параметров: " +
                            $"Значение {name_prop}: {name ?? "<пусто>"}; значение {description_prop}: {description ?? "<пусто>"}.");
                    }
                }
                else 
                    Output.Print($"Элемент «{saved_item.DisplayName}» уже присутствует в Quantification;");
            }
            GlobalPreferences.state = 0;
        }

        public static void UpdateQuantification(string folder_name, bool remove_empty)
        {
            Output.PrintHeader("KPLN Extention : Наполнение элементов каталога Quantification");
            if (remove_empty)
            { 
                DBTools.ClearSQLObjects();           
            }
            
            List<SavedItem> saved_items = APITools.GetSavedItems(folder_name);
            GlobalPreferences.Update();
            foreach (SavedItem saved_item in saved_items)
            {
                SelectionSet selectionSet = saved_item as SelectionSet;
                if (selectionSet.GetSelectedItems().Count == 0)
                {
                    Output.PrintAlert($"Поисковый набор «{saved_item.DisplayName}» не содержит элементов в модели");
                    continue;
                }

                DBItemGroup folder = DBTools.GetParentByNameContains(saved_item.DisplayName);
                if (folder != null)
                {
                    List<DBObject> objects = DBTools.GetObjectsByParentId(folder.Id);
                    ItemsSearch item_search = APITools.CreateNewSearch(saved_item);
                    foreach (Guid search_guid in item_search.Guids)
                    {
                        if (!DBTools.ItemInList(objects, search_guid))
                        {
                            string wbs = DBTools.GetNextWbs(objects);
                            //DBTools.DoTakeoff(folder, search_guid, "1");
                            //DBTools.GPInsertObject(folder, saved_item, objects, "1");
                        }
                    }

                }
                else 
                    Output.PrintAlert($"Для объекта «{saved_item.DisplayName}» не найдено родительского элемента;");
            }
            GlobalPreferences.state = 0;
        }

        public static void CreateSelectionSets(string project_name, string by_category, string by_property)
        {
            Output.PrintHeader("KPLN Extention : Создание поисковых наборов");
            string major_folder_name;
            if (project_name == null)
            {
                major_folder_name = string.Format("{0} ({1})", GlobalPreferences.GetDocumentName(), Guid.NewGuid().ToString());
            }
            else 
            {
                major_folder_name = project_name;
            }
            FolderItem major_folder = APITools.CreateNewFolder(major_folder_name);
            List<ModelItem> elements = DBTools.GetAllElements(project_name);
            List<string> values = APITools.GetUniqValuesOfDataProperty(elements, by_category, by_property);
            HashSet<string> split_values = APITools.GetSplittedValues(values, '.');

            List<FolderItem> folders = new List<FolderItem>();
            foreach (string value in split_values)
            {
                folders.Add(APITools.CreateNewFolder(value));
            }
            foreach (FolderItem folder_1 in folders)
            {
                if (!APITools.MoveFolderToParent(folder_1, folders))
                {
                    APITools.MoveFolderToFolder(folder_1, major_folder);
                }
            }
            foreach (string value in values)
            {
                SelectionSet selection_set = APITools.GetSelectionSetByProperty(by_category, by_property, value, value);
                if (!APITools.FindSelectionSetParent(selection_set, folders))
                {
                    APITools.MoveSelectionSetToFolder(selection_set, major_folder);
                }
                Output.PrintSuccess(string.Format("Поисковый набор «{0}» добавлен в наборы;", value));
            }
            GlobalPreferences.state = 0;
            //List<SavedItem> saved_items = APITools.GetSavedItems(project_name);
        }

        public static void MatchResources(string by_category, string by_parameter)
        {
            Output.PrintHeader("KPLN Extention : Сопоставление каталога ресурсов с каталогом элементов");
            GlobalPreferences.Update();
            List<DBItemGroup> empty_items = DBTools.MatchRes(GlobalPreferences.items, by_category, by_parameter);
            int count = 0;
            while (empty_items.ToArray().Length != 0)
            {
                count++;
                empty_items = DBTools.MatchRes(empty_items, by_category, by_parameter);
                if (count > 1000)
                {
                    break; 
                }
            }
            if (empty_items.ToArray().Length != 0) {
                List<string> elements = new List<string>();
                foreach (DBItemGroup item in empty_items)
                {
                    elements.Add("\t- " + item.GlobalWBS);
                }

                Output.PrintAlert(string.Format("Нерасчитанных элементов - {0};\n{1}", empty_items.ToArray().Length.ToString(), string.Join("\n", elements))); 
            }
            else 
                Output.PrintSuccess("Все элементы успешно расчитаны;");
        }
    }
}
