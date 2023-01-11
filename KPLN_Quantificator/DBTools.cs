using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Data;
using Autodesk.Navisworks.Api.Takeoff;
using Autodesk.Navisworks.Api.Clash;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KPLN_Quantificator.Forms;

namespace KPLN_Quantificator
{
    class DBTools
    {
        public static void GetClashes()
        {
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
            DocumentClash documentClash = document.GetClash();
            DocumentClashTests oDCT = documentClash.TestsData;

            foreach (ClashTest test in oDCT.Tests)
            {
                string name = test.DisplayName;
                ClashTestStatus status = test.Status;
                CommentCollection comments = test.Comments;
                foreach (SavedItem issue in test.Children)
                {
                    ClashResultGroup group = issue as ClashResultGroup;
                    if (null != group)
                    {
                        string groupName = group.DisplayName;
                        ClashResultStatus groupStatus = group.Status;
                        foreach (SavedItem issue1 in group.Children)
                        {
                            ClashResult rt1 = issue as ClashResult;

                        }

                    }
                    //
                    //ClashResult rt = issue as ClashResult;
//
                    //if (null != rt)
//
                    //    writeClashResult(rt, sw);

                }//Result

            } //Test
        }
        public static string StringToLatin(string value)
        {
            string dict_latin = "HKOPCMTEAXBcaxeopk";
            string dict_cyrilic = "НКОРСМТЕАХВсахеорк";
            string result = "";
            foreach (char a in value)
            {
                if (dict_cyrilic.Contains(a))
                {
                    int count = 0;
                    foreach (char b in dict_cyrilic)
                    {
                        if (b == a)
                        {
                            result += dict_latin[count];
                            break;
                        }
                        count++;
                    }
                }
                else
                {
                    result += a;
                }
            }
            return result;
        }
        public static long GetStepId(DBItemGroup item)
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = "SELECT ItemId FROM TK_Step";
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int v = 0;
                        while (true)
                        {
                            v++;
                            try
                            {
                                long step_id = reader.GetInt64(0);
                                reader.Read();
                                if (step_id == item.Id)
                                { return step_id; }
                            }
                            catch (Exception)
                            {
                                break;
                            }
                            if (v > 1000000)
                            {
                                break;
                            }

                        }
                    }
                }
            }
            return -1;
        }
        public static bool StepContainsStepResourse(long step_id)
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = "SELECT StepId FROM TK_StepResource";
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int v = 0;
                        while (true)
                        {
                            v++;
                            try
                            {
                                long resource_step_id = reader.GetInt64(0);
                                reader.Read();
                                if (resource_step_id == step_id)
                                { return true; }
                            }
                            catch (Exception)
                            {
                                break;
                            }
                            if (v > 1000000)
                            {
                                break;
                            }

                        }
                    }
                }
            }
            return false;
        }
        public static bool ItemHasResource(DBItemGroup item)
        {
            long step_id = GetStepId(item);
            if (step_id != -1)
            {
                return StepContainsStepResourse(step_id);
            }
            return false;
        }
        public static DBResourceGroup GetResourceByRBS(string RBS)
        {
            foreach (DBResourceGroup resourse in GlobalPreferences.resources)
            {
                if (StringToLatin(resourse.GlobalRBS) == StringToLatin(RBS))
                {
                    return resourse;
                }
            }
            foreach (DBResourceGroup resourse in GlobalPreferences.resources)
            {
                if (StringToLatin(resourse.GlobalRBS) == StringToLatin(APITools.GetSingleSplittedValue(RBS, '.')))
                {
                    return resourse;
                }
            }
            return null;
        }
        public static List<DBItemGroup> MatchRes(List<DBItemGroup> items, string by_category, string by_parameter)
        {
            List<DBItemGroup> empty_items = new List<DBItemGroup>();
            foreach (DBItemGroup item in items)
            {
                if (!ItemHasResource(item))
                {
                    List<DBObject> objects = new List<DBObject>();
                    foreach (DBObject obj in GlobalPreferences.objects)
                    {
                        if (obj.ParentId == item.Id)
                        {
                            objects.Add(obj);
                        }
                    }
                    if (objects.Count != 0)
                    {
                        string RBS = GetItemRBS(objects, by_category, by_parameter);
                        if (RBS == null) 
                        { 
                            RBS = GetItemParentRBS(objects, by_category, by_parameter);
                        }
                        if (RBS != null)
                        {
                            try
                            {
                                DBResourceGroup resource_target = GetResourceByRBS(RBS);
                                if (resource_target != null)
                                {
                                    int step_id = Insert_STEP(item);
                                    Insert_STEPRESOURCE(step_id, resource_target.Id);
                                }
                                else { Output.PrintAlert(string.Format("Для элемента {0} не найдено доступных ресурсов по значению «{1}»;", item.GlobalWBS, RBS)); }
                            }
                            catch (Exception)
                            {
                                empty_items.Add(item);
                            }
                        }
                    }
                }
            }
            return empty_items;
        }
        public static void GetResources()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = "SELECT ID, Parent, RBS FROM TK_Resource";
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int v = 0;
                        while (true)
                        {
                            v++;
                            try
                            {
                                DBResourceGroup item = new DBResourceGroup();
                                item.Id = reader.GetInt64(0);
                                try
                                {
                                    item.ParentId = reader.GetInt64(1);
                                }
                                catch (Exception)
                                {
                                    item.ParentId = -1;
                                }
                                item.RBS = reader.GetString(2);
                                item.Parent = null;
                                item.GlobalRBS = null;
                                GlobalPreferences.resources.Add(item);
                                reader.Read();
                            }
                            catch (Exception)
                            {
                                break;
                            }
                            if (v > 1000000)
                            {
                                break;
                            }

                        }
                    }
                }
                foreach (DBResourceGroup item in GlobalPreferences.resources)
                {
                    item.FindItemParent(GlobalPreferences.resource_groups);
                }
                foreach (DBResourceGroup item in GlobalPreferences.resources)
                {
                    item.ParseGlobalRBS(item.Parent.GlobalRBS);
                }
            }
        }
        public static void GetResourceGroups()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = "SELECT ID, Parent, RBS FROM TK_ResourceGroup";
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int v = 0;
                        while (true)
                        {
                            v++;
                            try
                            {
                                DBResourceGroup item = new DBResourceGroup();
                                item.Id = reader.GetInt64(0);
                                try
                                {
                                    item.ParentId = reader.GetInt64(1);
                                }
                                catch (Exception)
                                {
                                    item.ParentId = -1;
                                }
                                item.RBS = reader.GetString(2);
                                item.Parent = null;
                                item.GlobalRBS = null;
                                GlobalPreferences.resource_groups.Add(item);
                                reader.Read();
                            }
                            catch (Exception)
                            {
                                break;
                            }
                            if (v > 1000000)
                            {
                                break;
                            }

                        }
                    }
                }
                foreach (DBResourceGroup item in GlobalPreferences.resource_groups)
                {
                    item.FindParent(GlobalPreferences.resource_groups);
                }
                foreach (DBResourceGroup item in GlobalPreferences.resource_groups)
                {
                    if (item.Parent == null)
                    {
                        item.ParseGlobalRBS(null);
                    }
                }
            }
        }
        public static List<DBItemGroup> GetItems()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = "SELECT ID, Parent, WBS, Name FROM TK_Item";
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        int v = 0;
                        while (true)
                        {
                            v++;
                            try
                            {
                                DBItemGroup item = new DBItemGroup();
                                item.Id = reader.GetInt64(0);
                                try
                                {
                                    item.ParentId = reader.GetInt64(1);
                                }
                                catch (Exception)
                                {
                                    item.ParentId = -1;
                                }
                                item.WBS = reader.GetString(2);
                                item.Parent = null;
                                item.Name = reader.GetString(3);
                                item.GlobalWBS = null;
                                GlobalPreferences.items.Add(item);
                                reader.Read();
                            }
                            catch (Exception)
                            {
                                break;
                            }
                            if (v > 1000000)
                            {
                                break;
                            }

                        }
                    }
                }
                foreach (DBItemGroup item in GlobalPreferences.items)
                {
                    item.FindItemParent(GlobalPreferences.item_groups);
                }
                foreach (DBItemGroup item in GlobalPreferences.items)
                {
                    item.ParseGlobalWBS(item.Parent.GlobalWBS);
                }
                return GlobalPreferences.items;
            }
        }
        public static List<DBObject> GetObjects()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                try
                {
                    //
                    cmd.CommandText = "SELECT ID, Parent, WBS, ModelItemId FROM TK_Object";
                    using (NavisWorksDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            int v = 0;
                            while (true)
                            {
                                v++;
                                try
                                {
                                    DBObject item = new DBObject();
                                    item.Id = reader.GetInt64(0);
                                    try
                                    {
                                        item.ParentId = reader.GetInt64(1);
                                    }
                                    catch (Exception)
                                    {
                                        item.ParentId = -1;
                                    }
                                    try
                                    {
                                        item.WBS = reader.GetString(2);
                                    }
                                    catch (Exception)
                                    {
                                        item.WBS = null;
                                    }
                                    item.Guid = reader.GetGuid(3);
                                    item.Parent = null;
                                    item.GlobalWBS = null;
                                    GlobalPreferences.objects.Add(item);
                                    reader.Read();
                                }
                                catch (Exception)
                                {
                                    break;
                                }
                                if (v > 1000000)
                                {
                                    break;
                                }

                            }
                        }
                    }
                    foreach (DBObject item in GlobalPreferences.objects)
                    {
                        item.FindParent(GlobalPreferences.items);
                    }
                    List<DBItemGroup> parent_items = new List<DBItemGroup>();
                    foreach (DBItemGroup item in GlobalPreferences.item_groups)
                    {
                        if (item.Parent == null)
                        {
                            parent_items.Add(item);
                        }
                    }
                }
                catch (Exception e)
                {
                    Output.PrintError(e);
                }
                return GlobalPreferences.objects;
            }
        }
        public static List<DBItemGroup> GetItemGroups()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                try
                {
                    cmd.CommandText = "SELECT ID, Parent, WBS, Name FROM TK_ItemGroup";
                    using (NavisWorksDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            int v = 0;
                            while (true)
                            {
                                v++;
                                try
                                {
                                    DBItemGroup item = new DBItemGroup();
                                    item.Id = reader.GetInt64(0);
                                    try
                                    {
                                        item.ParentId = reader.GetInt64(1);
                                    }
                                    catch (Exception)
                                    {
                                        item.ParentId = -1;
                                    }
                                    item.WBS = reader.GetString(2);
                                    item.Parent = null;
                                    item.Name = reader.GetString(3);
                                    item.GlobalWBS = null;
                                    GlobalPreferences.item_groups.Add(item);
                                    reader.Read();
                                }
                                catch (Exception)
                                {
                                    break;
                                }
                                if (v > 1000000)
                                {
                                    break;
                                }

                            }
                        }
                    }
                    foreach (DBItemGroup item in GlobalPreferences.item_groups)
                    {
                        item.FindParent(GlobalPreferences.item_groups);
                    }
                    List<DBItemGroup> parent_items = new List<DBItemGroup>();
                    foreach (DBItemGroup item in GlobalPreferences.item_groups)
                    {
                        if (item.Parent == null)
                        {
                            parent_items.Add(item);
                        }
                    }
                    foreach (DBItemGroup item in parent_items)
                    {
                        item.ParseGlobalWBS(null);
                    }
                }
                catch (Exception) { }
                return GlobalPreferences.item_groups;
            }
        }
        public static DBItemGroup GetItemByWBS(string wbs)
        {
            foreach (DBItemGroup item in GlobalPreferences.items)
            {
                if (item.GlobalWBS.ToString() == wbs.ToString())
                {
                    return item;
                }
            }
            return null;
        }
        public static List<DBObject> GetObjectsByParentId(long parent_id)
        {
            List<DBObject> objects = new List<DBObject>();
            foreach (DBObject item in GlobalPreferences.objects)
            {
                if (item.ParentId == parent_id)
                {
                    objects.Add(item);
                }
            }
            return objects;
        }
        public static bool ItemInList(List<DBObject> objects, Guid guid)
        {
            foreach (DBObject obj in objects)
            {
                if (obj.Guid == guid)
                {
                    return true;
                }
            }
            return false;
        }

        public static bool WbsInList(List<DBObject> objects, string wbs)
        {
            foreach (DBObject obj in objects)
            {
                if (obj.WBS == wbs)
                {
                    return true;
                }
            }
            return false;
        }

        public static string GetNextWbs(List<DBObject> objects)
        {
            int wbs = 1;
            while (true)
            {
                if (!WbsInList(objects, wbs.ToString()))
                {
                    return wbs.ToString();
                }
                wbs++;
            }
        }

        public static long GetLastInsertRowId()
        {
            DocumentTakeoff doc_take_off = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = doc_take_off.Database.Value.CreateCommand())
            {
                cmd.CommandText = "select last_insert_rowid()";
                using (NavisWorksDataReader dataReader = cmd.ExecuteReader())
                {
                    long last_insert_id = -1;
                    if (dataReader.Read())
                    {
                        long.TryParse(dataReader[0].ToString(), out last_insert_id);
                    }
                    return last_insert_id;
                }
            }
        }

        public static void InsertItemGroup(long? parent, string name, string description, string wbs)
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            string sql = "INSERT INTO TK_ItemGroup(parent, name, description, wbs) VALUES(@parent, @name, @description, @wbs)";
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    NavisworksParameter p = cmd.CreateParameter();
                    p.ParameterName = "@parent";
                    if (parent.HasValue)
                    {
                        p.Value = parent.Value;
                    }
                    else
                    {
                        p.Value = null;
                    }

                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@name";
                    p.Value = name;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@description";
                    p.Value = description;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@wbs";
                    p.Value = wbs;
                    cmd.Parameters.Add(p);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();
            }
        }

        public static int GetLastWbs(long? parent)
        {
            List<int> wbsList = new List<int>();
            
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    cmd.CommandText = $"SELECT wbs FROM TK_Item WHERE parent={parent}";
                    using (NavisWorksDataReader dataReader = cmd.ExecuteReader())
                    {
                        while (dataReader.Read())
                        {
                            if (int.TryParse(dataReader.GetString(0), out int wbs))
                            {
                                wbsList.Add(wbs);

                            }
                            else
                                throw new Exception("Конечный коды WBS (Структура рабочего распределения) должен быть только циферным");
                        }
                    }
                }
                trans.Commit();
            }
            if (wbsList.Count == 0)
                return 0;
            else
                return wbsList.Max();
        }

        public static long InsertItem(long? parent, string name, string description, string wbs)
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            string sql = "INSERT INTO TK_ITEM(parent, name, description, wbs, color, transparency, linethickness, countsymbol, countsize) VALUES(@parent, @name, @description,@wbs, @color, @transparency, @linethickness, @countsymbol, @countsize)";
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    NavisworksParameter p = cmd.CreateParameter();
                    p.ParameterName = "@parent";
                    if (parent.HasValue)
                    {
                        p.Value = parent.Value;
                    }
                    else
                    {
                        p.Value = null;
                    }

                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@name";
                    p.Value = name;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@description";
                    p.Value = description;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@wbs";
                    p.Value = wbs;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@color";
                    p.Value = 24;// 24, 0.5, 0.1, 1, 2
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@transparency";
                    p.Value = 0.5;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@linethickness";
                    p.Value = 0.1;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@countsymbol";
                    p.Value = 1;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@countsize";
                    p.Value = 2;
                    cmd.Parameters.Add(p);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();
            }
            return GetLastInsertRowId();
        }
        public static int GetLastTableId(string table)
        {
            int id = 1;
            HashSet<int> ids = new HashSet<int>();
            int step = 0;
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
            {
                cmd.CommandText = string.Format("SELECT ID FROM {0}", table);
                using (NavisWorksDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (step < 1000000)
                        {
                            try
                            {
                                ids.Add(int.Parse(reader.GetValue(0).ToString()));
                                reader.Read();
                            }
                            catch (Exception)
                            {
                                break;
                            }
                        }
                    }
                }
            }
            while (ids.Contains(id))
            {
                id++;
            }
            return id;

        }
        public static int Insert_STEP(DBItemGroup item)
        {
            int last_id = GetLastTableId("TK_Step");
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            string sql = "INSERT INTO TK_Step(Id, name, ItemId) VALUES(@Id, @name, @ItemId)";
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    NavisworksParameter p = cmd.CreateParameter();
                    p.ParameterName = "@Id";
                    p.Value = last_id;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@name";
                    p.Value = "New step";
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@ItemId";
                    p.Value = item.Id;
                    cmd.Parameters.Add(p);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();
            }
            return last_id;
        }
        public static void Insert_STEPRESOURCE(long? object_step_id, long? step_resource_id)
        {
            int last_id = DBTools.GetLastTableId("TK_StepResource");
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            string sql = "INSERT INTO TK_StepResource(Id, StepId, ResourceId) VALUES(@Id, @StepId, @ResourceId)";
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    NavisworksParameter p = cmd.CreateParameter();
                    p.ParameterName = "@Id";
                    p.Value = last_id;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@StepId";
                    p.Value = object_step_id.Value;
                    cmd.Parameters.Add(p);

                    p = cmd.CreateParameter();
                    p.ParameterName = "@ResourceId";
                    p.Value = step_resource_id.Value;
                    cmd.Parameters.Add(p);

                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                trans.Commit();
            }
        }
        public static void CreateFolderStructure(QFolder folder, int parent_index)
        {
            folder.Index = GlobalPreferences.GetGroupIndex();
            folder.Parent = parent_index;
            if (folder.Parent == -1)
            {
                InsertItemGroup(null, folder.Name, "", folder.Index.ToString());
            }
            else
            {
                InsertItemGroup(folder.Parent, folder.Name, "", folder.Index.ToString());
            }
            if (folder.SearchList.Count > 0)
            {
                int local_wbs = 1;
                foreach (ItemsSearch items_search in folder.SearchList)
                {
                    int item_index = GlobalPreferences.GetItemIndex();
                    InsertItem(folder.Index, items_search.Name, folder.Name, local_wbs.ToString()); //создание записей
                    local_wbs += 1;
                }
            }
            if (folder.Folders.Count > 0)
            {
                foreach (QFolder subfolder in folder.Folders)
                {
                    CreateFolderStructure(subfolder, folder.Index);
                }
            }


        }
        public static void ClearSQLObjects()
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {
                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM TK_Object";
                    cmd.ExecuteReader();
                }
                trans.Commit();
            }
        }
        public static string GetItemRBS(List<DBObject> objects, string by_category, string by_property)
        {
            foreach (DBObject obj in objects)
            {
                if (obj.Guid.ToString() != new Guid().ToString())
                {
                    try
                    {
                        ModelItem item = Autodesk.Navisworks.Api.Application.MainDocument.Models.RootItemDescendantsAndSelf.WhereInstanceGuid(obj.Guid).ToList()[0];
                        foreach (PropertyCategory cat in item.PropertyCategories)
                        {
                            if (cat.DisplayName == by_category)
                            {
                                foreach (DataProperty prop in cat.Properties)
                                {
                                    if (prop.DisplayName == by_property)
                                    {
                                        try
                                        {
                                            string value = prop.Value.ToDisplayString().ToString();
                                            if (value != "")
                                            {
                                                return value;
                                            }
                                            
                                        }
                                        catch (Exception) { }
                                    }
                                }
                            }
                        }

                    }
                    catch (Exception e) { Output.PrintError(e); }
                }
            }
            return null;
        }
        public static string GetItemSubParentRBS(ModelItem item, string by_category, string by_property)
        {
            try
            {
                foreach (PropertyCategory cat in item.Parent.PropertyCategories)
                {
                    if (cat.DisplayName == by_category)
                    {
                        foreach (DataProperty prop in cat.Properties)
                        {
                            if (prop.DisplayName == by_property)
                            {
                                string value = prop.Value.ToDisplayString().ToString();
                                if (value != "")
                                {
                                    return value;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e) { Output.PrintError(e); }
            try
            {
                string sub_value = GetItemSubParentRBS(item.Parent, by_category, by_property);
                if (sub_value != null)
                {
                    return sub_value;
                }
            }
            catch (Exception e) { Output.PrintError(e); }
            return null;
        }
        public static string GetItemParentRBS(List<DBObject> objects, string by_category, string by_property)
        {
            foreach (DBObject obj in objects)
            {
                if (obj.Guid.ToString() != new Guid().ToString())
                {
                    try
                    {
                        ModelItem item = Autodesk.Navisworks.Api.Application.MainDocument.Models.RootItemDescendantsAndSelf.WhereInstanceGuid(obj.Guid).ToList()[0];
                        string value = GetItemSubParentRBS(item, by_category, by_property);
                        if (value != null)
                        {
                            return value;
                        }
                    }
                    catch (Exception e) { Output.PrintError(e); }
                }
            }
            return null;
        }
        
        public static void DoTakeoff(long itemRowId, Guid model_item_guid, string wbs_number) //вызываем из Main 
        {
            DocumentTakeoff docTakeoff = Autodesk.Navisworks.Api.Application.MainDocument.GetTakeoff();
            ModelItem item = Autodesk.Navisworks.Api.Application.MainDocument.Models.RootItemDescendantsAndSelf.WhereInstanceGuid(model_item_guid).FirstOrDefault();
            
            using (NavisworksTransaction trans = docTakeoff.Database.BeginTransaction(DatabaseChangedAction.Edited))
            {

                //using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                //{
                //    cmd.CommandText = $"SELECT RowId FROM TK_Item WHERE parent={itemGroup.Id}";
                //    using (NavisWorksDataReader dataReader = cmd.ExecuteReader())
                //    {
                //        while (dataReader.Read())
                //        {
                //            Output.Print($"{dataReader.GetInt32(0)}");
                //        }
                //    }
                //}

                docTakeoff.Objects.InsertModelItemTakeoff(itemRowId, item);
                long lastId = GetLastInsertRowId();

                using (NavisworksCommand cmd = docTakeoff.Database.Value.CreateCommand())
                {
                    cmd.CommandText = $"UPDATE TK_OBJECT SET wbs = {wbs_number} WHERE id = {lastId}";
                    cmd.ExecuteNonQuery();
                }

                trans.Commit();
            }
        }
        public static List<ModelItem> GetAllElements(string model_name)
        {
            try
            {
                List<ModelItem> elements = new List<ModelItem>();
                if (model_name != null)
                {
                    ModelItem nwc_model = GetModelFile(model_name);
                    if (nwc_model != null)
                    {
                        foreach (ModelItem item in GetAllSubElements(nwc_model))
                        {
                            { elements.Add(item); }
                        } 
                    }
                    return elements;
                }
                else
                {
                    foreach (ModelItem model_item in GetAllModels())
                    {
                        if (model_item != null)
                        {
                            foreach (ModelItem item in GetAllSubElements(model_item))
                            {
                                {
                                    elements.Add(item);
                                }
                            }
                        }
                    }
                    return elements;
                }
            }
            catch (Exception)
            {
                return null;
            }

        }
        public static List<ModelItem> GetAllSubElements(ModelItem model)
        {
            try
            {
                List<ModelItem> elements = new List<ModelItem>();

                foreach (ModelItem item in model.Children)
                {
                    if (item.IsLayer || item.IsCollection)
                    {
                        foreach (ModelItem it in GetAllSubElements(item))
                        {
                            elements.Add(it);
                        }
                    }
                    else
                    {
                        elements.Add(item);
                    }
                }

                return elements;
            }
            catch (Exception)
            {
                return null;
            }

        }
        public static List<ModelItem> GetAllModels()
        {
            List<ModelItem> models = new List<ModelItem>();
            try
            {
                foreach (Model model in Autodesk.Navisworks.Api.Application.ActiveDocument.Models)
                {
                    foreach (ModelItem submodel in model.RootItem.Children)
                    {
                        models.Add(submodel);
                    }
                }
                return models;
            }
            catch (Exception) { }
            return models;
        }
        public static ModelItem GetModelFile(string model_name)
        {
            try
            {
                foreach (Model model in Autodesk.Navisworks.Api.Application.ActiveDocument.Models)
                {
                    foreach (ModelItem submodel in model.RootItem.Children)
                    {
                        if (model_name.Contains(submodel.DisplayName) || model_name == null)
                        {
                            return submodel;
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
        
        public static DBItemGroup GetParentByWBS(string wbs)
        {
            foreach (DBItemGroup item_group in GlobalPreferences.item_groups)
            {
                if (item_group.GlobalWBS == APITools.GetSingleSplittedValue(wbs, '.'))
                {
                    return item_group;
                }
            }
            return null;
        }

        public static DBItemGroup GetParentByNameContains(string savedItemName)
        {
            string[] splitedName = savedItemName.Split(new string[] { "<=>" }, StringSplitOptions.None);
            foreach (DBItemGroup item_group in GlobalPreferences.item_groups)
            {
                foreach(string name in splitedName)
                {
                    if (name.Trim().Equals(item_group.GlobalWBS))
                    {
                        //Output.PrintHeader($"{name.Trim()} : {item_group.GlobalWBS} / {item_group.WBS}");
                        return item_group;
                    }
                }
            }
            return null;
        }

        public static void GPInsertItem(DBItemGroup parent_folder, SavedItem saved_item, string local_wbs)
        {
            DBItemGroup new_item = new DBItemGroup();
            new_item.GlobalWBS = saved_item.DisplayName;
            new_item.ParentId = parent_folder.Id;
            new_item.Parent = parent_folder;
            new_item.WBS = local_wbs;
            GlobalPreferences.items.Add(new_item);
        }
        
        public static void GPInsertObject(DBItemGroup parent_folder, SavedItem saved_item, List<DBObject> object_list, string local_wbs)
        {
            DBObject new_item = new DBObject();
            new_item.GlobalWBS = parent_folder.GlobalWBS + "." + local_wbs;
            new_item.ParentId = parent_folder.Id;
            new_item.Parent = parent_folder;
            new_item.WBS = local_wbs;
            object_list.Add(new_item);
        }
    }
    
    public class DBObject
    {
        public long Id { get; set; }
        
        public long ParentId { get; set; }
        
        public string WBS { get; set; }
        
        public string GlobalWBS { get; set; }
        
        public Guid Guid { get; set; }
        
        public DBItemGroup Parent { get; set; }

        public void FindParent(List<DBItemGroup> item_list)
        {
            foreach (DBItemGroup item in item_list)
            {
                if (item.Id == ParentId)
                {
                    Parent = item;
                    GlobalWBS = Parent.GlobalWBS + "." + WBS;
                    return;
                }
            }
            Parent = null;
        }

    }
    
    public class DBItemGroup
    {
        
        public long Id { get; set; }

        public string Name { get; set; }

        public long ParentId { get; set; }
        
        public string WBS { get; set; }
        
        public string GlobalWBS { get; set; }
        
        public DBItemGroup Parent { get; set; }
        
        public List<DBItemGroup> Child = new List<DBItemGroup>();

        public void ParseGlobalWBS(string wbs)
        {
            if (wbs == null)
            {
                GlobalWBS = WBS;
                if (Child.Count > 0)
                {
                    foreach (DBItemGroup child_item in Child)
                    {
                        child_item.ParseGlobalWBS(WBS);
                    }
                }
            }
            else 
            {
                string new_wbs = string.Format("{0}.{1}", wbs, WBS);
                GlobalWBS = new_wbs;
                if (Child.Count > 0)
                {
                    foreach (DBItemGroup child_item in Child)
                    {
                        child_item.ParseGlobalWBS(new_wbs);
                    }
                }
            }
        }
        
        public void FindItemParent(List<DBItemGroup> item_list)
        {
            foreach (DBItemGroup item in item_list)
            {
                if (item.Id == ParentId)
                {
                    Parent = item;
                    return;
                }
            }
            Parent = null;
        }
        
        public void FindParent(List<DBItemGroup> item_list)
        {
            foreach (DBItemGroup item in item_list)
            {
                if (item.Id == ParentId)
                {
                    Parent = item;
                    item.Child.Add(this);
                    return;
                }
            }
            Parent = null;
        }
    }
    public class DBResourceGroup
    {
        public long Id { get; set; }
        
        public long ParentId { get; set; }
        
        public string RBS { get; set; }
        
        public string GlobalRBS { get; set; }
        
        public DBResourceGroup Parent { get; set; }
        
        public List<DBResourceGroup> Child = new List<DBResourceGroup>();

        public void ParseGlobalRBS(string rbs)
        {
            if (rbs == null)
            {
                GlobalRBS = RBS;
                if (Child.Count > 0)
                {
                    foreach (DBResourceGroup child_item in Child)
                    {
                        child_item.ParseGlobalRBS(RBS);
                    }
                }
            }
            else
            {
                string new_rbs = string.Format("{0}.{1}", rbs, RBS);
                GlobalRBS = new_rbs;
                if (Child.Count > 0)
                {
                    foreach (DBResourceGroup child_item in Child)
                    {
                        child_item.ParseGlobalRBS(new_rbs);
                    }
                }
            }
        }
        public void FindItemParent(List<DBResourceGroup> item_list)
        {
            foreach (DBResourceGroup item in item_list)
            {
                if (item.Id == ParentId)
                {
                    Parent = item;
                    return;
                }
            }
            Parent = null;
        }
        public void FindParent(List<DBResourceGroup> item_list)
        {
            foreach (DBResourceGroup item in item_list)
            {
                if (item.Id == ParentId)
                {
                    Parent = item;
                    item.Child.Add(this);
                    return;
                }
            }
            Parent = null;
        }
    }
}
