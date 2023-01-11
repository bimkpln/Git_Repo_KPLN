using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Quantificator
{
    public class SavedItemCategory
    {
        public string DisplayName { get; set; }
        public HashSet<string> HSParameters = new HashSet<string>();
    }
    public class QDataRow
    { 
        public long ItemId { get; set; }
        public Guid ModelGuid { get; set; }
    }

    class GlobalPreferences
    {
        public static int state = 0;
        private static int item_index = 0;
        private static int group_index = 0;
        public static List<DBItemGroup> item_groups = new List<DBItemGroup>();
        public static List<DBItemGroup> items = new List<DBItemGroup>();
        public static List<DBObject> objects = new List<DBObject>();
        public static List<DBResourceGroup> resource_groups = new List<DBResourceGroup>();
        public static List<DBResourceGroup> resources = new List<DBResourceGroup>();
        public static List<SavedItemCategory> project_categories = new List<SavedItemCategory>();
        public static string GetDocumentName()
        {
            List<string> parts = Autodesk.Navisworks.Api.Application.ActiveDocument.FileName.Split('\\').ToList();
            return parts[parts.Count-1];
        }
        public static void Update()
        {
            item_index = 0;
            group_index = 0;
            item_groups = new List<DBItemGroup>();
            items = new List<DBItemGroup>();
            objects = new List<DBObject>();
            resource_groups = new List<DBResourceGroup>();
            resources = new List<DBResourceGroup>();
            DBTools.GetItemGroups();
            DBTools.GetItems();
            DBTools.GetObjects();
            DBTools.GetResourceGroups();
            DBTools.GetResources();
        }
        public static int GetItemIndex()
        {
            item_index++;
            return item_index;
        }
        public static int GetGroupIndex()
        {
            group_index++;
            return group_index;
        }
        public static List<SavedItemCategory> GetCategories(List<ModelItem> elements)
        {
            HashSet<string> uniq_categories = new HashSet<string>();
            List<SavedItemCategory> categories = new List<SavedItemCategory>();

            foreach (ModelItem item in elements)
            {
                foreach (PropertyCategory prop_cat in item.PropertyCategories)
                {
                    uniq_categories.Add(prop_cat.DisplayName);

                }
            }
            foreach (string category in uniq_categories)
            {
                SavedItemCategory new_cat = new SavedItemCategory();
                new_cat.DisplayName = category;
                categories.Add(new_cat);
            }
            foreach (ModelItem item in elements)
            {
                foreach (PropertyCategory prop_cat in item.PropertyCategories)
                {
                    foreach (SavedItemCategory uniq_cat in categories)
                    {
                        if (uniq_cat.DisplayName == prop_cat.DisplayName)
                        {
                            foreach (DataProperty data_prop in prop_cat.Properties)
                            {
                                uniq_cat.HSParameters.Add(data_prop.DisplayName);
                            }
                        }
                    }
                }
            }
            //categories.Sort();
            return categories;
        }
    }
}
