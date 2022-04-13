using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_Scoper.Common;
using KPLN_Scoper.Tools;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;
using static KPLN_Loader.Output.Output;

namespace KPLN_Scoper
{
    public class Module : IExternalModule
    {
        private string _versionName { get; set; }
        public Result Close()
        {
            try
            {
                OverwriteIni(_versionName);
            }
            catch (Exception) { }
            try
            {
                ActivityManager.Synchronize(null);
                ActivityManager.Destroy();
            }
            catch (Exception)
            { }
            return Result.Succeeded;
        }
        public Result Execute(UIControlledApplication application, string tabName)
        {
            try
            {
                try
                {
                    _versionName = application.ControlledApplication.VersionNumber;
                    OverwriteIni(_versionName);
                }
                catch (Exception) { }
                try
                {
                    UpdateAllDocumentInfo();
                }
                catch (Exception) { }
                application.ControlledApplication.DocumentOpened += OnDocumentOpened;
                application.ControlledApplication.DocumentSynchronizedWithCentral += OnDocumentSynchronized;
                application.ViewActivated += OnViewActivated;
                application.ControlledApplication.DocumentChanged += OnDocumentChanged;
                try
                {
                    ActivityManager.Run();
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
                return Result.Succeeded;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
        private void OnViewActivated(object sender, ViewActivatedEventArgs args)
        {
            try
            {
                if (args.Document.Title != null)
                {
                    ActivityManager.ActiveDocument = args.Document;
                }
            }
            catch (Exception)
            {
                ActivityManager.ActiveDocument = null;
            }
        }
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            try
            {
                if (ActivityManager.ActiveDocument != null)
                {
                    ActivityInfo info = new ActivityInfo(ActivityManager.ActiveDocument, Collections.BuiltInActivity.DocumentChanged);
                    ActivityManager.ActivityBag.Enqueue(info);
                }
            }
            catch (Exception) { }
        }
        private SQLProject GetProjectById(List<SQLProject> list, int id)
        {
            foreach (SQLProject project in list)
            {
                if (project.Id == id)
                {
                    return project;
                }
            }
            return null;
        }
        private SQLDepartment GetDepartmentById(List<SQLDepartment> list, int id)
        {
            foreach (SQLDepartment department in list)
            {
                if (department.Id == id)
                {
                    return department;
                }
            }
            return null;
        }
        private void UpdateAllDocumentInfo()
        {
            if (KPLN_Loader.Preferences.User.SystemName != "iperfilyev")
            { return; }
            List<SQLProject> projects = GetProjects();
            List<SQLDepartment> departments = GetDepartments();
            foreach (SQLDocument doc in GetDocuments())
            {
                if (new FileInfo(doc.Path).Exists)
                {
                    try
                    {
                        if (doc.Department != -1 && doc.Project != -1)
                        {
                            SQLProject pickedProject = GetProjectById(projects, doc.Project);
                            SQLDepartment pickedDepartment = GetDepartmentById(departments, doc.Department);
                            if (pickedProject == null || pickedDepartment == null) { continue; }
                            string code = string.Format("{0}_{1}", pickedProject.Code, pickedDepartment.Code);
                            if (doc.Code == "NONE" || doc.Code != code)
                            {
                                SQLiteConnection sql = new SQLiteConnection();
                                try
                                {
                                    sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                                    sql.Open();
                                    SQLiteCommand cmd = new SQLiteCommand(string.Format("UPDATE Documents SET Code = '{0}' WHERE Id = {1}", code, doc.Id), sql);
                                    cmd.ExecuteNonQuery();
                                    sql.Close();
                                }
                                catch (Exception) { }
                                {
                                    try
                                    {
                                        sql.Close();
                                    }
                                    catch (Exception) { }
                                }
                            }
                        }
                    }
                    catch (Exception) { }
                }
                else
                {
                    try
                    {
                        SQLiteConnection db = new SQLiteConnection(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                        try
                        {
                            db.Open();
                            SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("DELETE FROM Documents WHERE Path='{0}'", doc.Path), db);
                            cmd_insert.ExecuteNonQuery();
                            db.Close();
                        }
                        catch (Exception)
                        {
                            db.Close();
                        }
                    }
                    catch (Exception) { }
                }
            }
        }
        /*
        private void UpdateRoomDictKeys(Document doc)
        {
            if (new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count == 0)
            {
                return;
            }
            HashSet<string> room_names = new HashSet<string>();
            HashSet<string> room_departments = new HashSet<string>();
            foreach (Room room in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
            {
                try
                {
                    room_names.Add(room.get_Parameter( BuiltInParameter.ROOM_NAME).AsString().ToLower());
                }
                catch (Exception)
                { }
                try
                {
                    string dep = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString().ToLower();
                    if (dep != string.Empty)
                    {
                        room_departments.Add(dep);
                    }
                }
                catch (Exception)
                { }
            }
            Thread t = new Thread(() =>
            {
                SQLiteConnection sql = new SQLiteConnection();
                try
                {
                    sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Dictionary.db;Version=3;");
                    sql.Open();
                    HashSet<string> stored_names = new HashSet<string>();
                    HashSet<string> stored_departments = new HashSet<string>();
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT KEY FROM DICT_ROOMS_NAMES", sql))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                stored_names.Add(rdr.GetString(0));
                            }
                        }
                    }
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT KEY FROM DICT_ROOMS_DEPARTMENTS", sql))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                stored_departments.Add(rdr.GetString(0));
                            }
                        }
                    }
                    foreach (string value in room_names)
                    {
                        if (!stored_names.Contains(value))
                        {
                            SQLiteCommand cmd_insert = new SQLiteCommand("INSERT INTO DICT_ROOMS_NAMES ([KEY], [TYPE]) VALUES(@KEY, @TYPE)", sql);
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KEY", Value = value });
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TYPE", Value = -1 });
                            cmd_insert.ExecuteNonQuery();
                        }
                    }
                    foreach (string value in room_departments)
                    {
                        if (!stored_departments.Contains(value))
                        {
                            SQLiteCommand cmd_insert = new SQLiteCommand("INSERT INTO DICT_ROOMS_DEPARTMENTS ([KEY], [TYPE]) VALUES(@KEY, @TYPE)", sql);
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KEY", Value = value });
                            cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TYPE", Value = -1 });
                            cmd_insert.ExecuteNonQuery();
                        }
                    }
                    sql.Close();
                }
                catch (Exception) { }
                {
                    try
                    {
                        sql.Close();
                    }
                    catch (Exception) { }
                }
            });
            t.Start();
        }
        */
        private void OnDocumentSynchronized(object sender, DocumentSynchronizedWithCentralEventArgs args)
        {
            try
            {
                if (ActivityManager.ActiveDocument != null)
                {
                    ActivityInfo info = new ActivityInfo(ActivityManager.ActiveDocument, Collections.BuiltInActivity.DocumentSynchronized);
                    ActivityManager.ActivityBag.Enqueue(info);
                }
            }
            catch (Exception) { }
            /*
            try
            {
                Autodesk.Revit.DB.Document document = args.Document;
                UpdateRoomDictKeys(document);
            }
            catch (Exception)
            { }
            */
            try
            { ActivityManager.Synchronize(null); }
            catch (Exception) { }
        }
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            try
            {
                if (ActivityManager.ActiveDocument != null)
                {
                    ActivityInfo info = new ActivityInfo(ActivityManager.ActiveDocument, Collections.BuiltInActivity.ActiveDocument);
                    ActivityManager.ActivityBag.Enqueue(info);
                }
            }
            catch (Exception) { }
            try
            {
                Autodesk.Revit.DB.Document document = args.Document;

                if (document.IsWorkshared && !document.IsDetached)
                {
                    string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(document.GetWorksharingCentralModelPath());
                    string name = path.Split(new string[] { @"\" }, StringSplitOptions.None).Last();
                    List<string> parts = new List<string>();
                    SQLProject pickedProject = null;
                    SQLDepartment pickedDepartment = null;
                    foreach (SQLProject prj in GetProjects())
                    {
                        if (prj.Keys.Contains('*'))
                        {
                            foreach (string part in prj.Keys.Split('*'))
                            {
                                if (path != string.Empty && path.Contains(part))
                                {
                                    pickedProject = prj;
                                }
                            }
                        }
                        else
                        {
                            if (path.Contains(prj.Keys))
                            {
                                pickedProject = prj;
                            }
                        }

                    }
                    foreach (SQLDepartment dep in GetDepartments())
                    {
                        if (name.Contains(dep.Code) || name.Contains(dep.CodeUs))
                        {
                            pickedDepartment = dep;
                        }
                    }
                    int department = -1;
                    int project = -1;
                    if (pickedProject != null)
                    {
                        project = pickedProject.Id;
                    }
                    if (pickedDepartment != null)
                    {
                        department = pickedDepartment.Id;
                    }
                    string code = "NONE";
                    if (pickedProject != null && pickedDepartment != null)
                    {
                        code = string.Format("{0}_{1}", pickedProject.Code, pickedDepartment.Code);
                    }
                    SQLDocument openedDocument = new SQLDocument(-1, path, name, department, project, code);
                    bool documentExist = false;
                    foreach (SQLDocument doc in GetDocuments())
                    {

                        if (doc.Path == openedDocument.Path)
                        {
                            documentExist = true;
                        }
                    }
                    if (!documentExist)
                    {
                        AddDocument(openedDocument);
                    }
                }
                foreach (RevitLinkInstance link in new FilteredElementCollector(document).OfClass(typeof(RevitLinkInstance)).WhereElementIsNotElementType().ToElements())
                {
                    try
                    {
                        Autodesk.Revit.DB.Document linkDocument = link.GetLinkDocument();
                        if (linkDocument.IsWorkshared)
                        {
                            string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(linkDocument.GetWorksharingCentralModelPath());
                            string name = path.Split(new string[] { @"\" }, StringSplitOptions.None).Last();
                            List<string> parts = new List<string>();
                            SQLProject pickedProject = null;
                            SQLDepartment pickedDepartment = null;
                            foreach (SQLProject prj in GetProjects())
                            {
                                if (prj.Keys.Contains('*'))
                                {
                                    foreach (string part in prj.Keys.Split('*'))
                                    {
                                        if (path != string.Empty && path.Contains(part))
                                        {
                                            pickedProject = prj;
                                        }
                                    }
                                }
                                else
                                {
                                    if (path.Contains(prj.Keys))
                                    {
                                        pickedProject = prj;
                                    }
                                }

                            }
                            foreach (SQLDepartment dep in GetDepartments())
                            {
                                if (name.Contains(dep.Code) || name.Contains(dep.CodeUs))
                                {
                                    pickedDepartment = dep;
                                }
                            }
                            int department = -1;
                            int project = -1;
                            if (pickedProject != null)
                            {
                                project = pickedProject.Id;
                            }
                            if (pickedDepartment != null)
                            {
                                department = pickedDepartment.Id;
                            }
                            string code = "NONE";
                            if (pickedProject != null && pickedDepartment != null)
                            {
                                code = string.Format("{0}_{1}", pickedProject.Code, pickedDepartment.Code);
                            }
                            SQLDocument scopeDocument = new SQLDocument(-1, path, name, department, project, code);
                            bool documentExist = false;
                            foreach (SQLDocument doc in GetDocuments())
                            {
                                if (doc.Path == scopeDocument.Path)
                                {
                                    documentExist = true;
                                }
                            }
                            if (!documentExist)
                            {
                                AddDocument(scopeDocument);
                            }
                        }
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
        }
        private bool AddDocument(SQLDocument document)
        {
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = sql.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO Documents ([Path], [Name], [Department], [Project], [Code]) VALUES (@Path, @Name, @Department, @Project, @Code)";
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName= "@Path", Value = document.Path });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Name", Value = document.Name });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Department", Value = document.Department });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Project", Value = document.Project });
                    cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Code", Value = document.Code });
                    cmd.ExecuteNonQuery();
                }
                sql.Close();
                return true;
            }
            catch (Exception) { }
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
                
            }
            return false;
        }
        private List<SQLDocument> GetDocuments()
        {
            List<SQLDocument> documents = new List<SQLDocument>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Documents", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetInt32(0) == -1) { continue; }
                            documents.Add(new SQLDocument(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetInt32(3), rdr.GetInt32(4), rdr.GetString(5)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception) { }
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return documents;
        }
        private bool IsCopy(string name)
        {
            List<string> parts = name.Split('.').ToList();
            if (parts.Count <= 2) { return false; }
            parts.RemoveAt(parts.Count - 1);
            if (parts.Last().Length == 4 && parts.Last().StartsWith("0")) { return true; }
            return false;
        }
        private string OnlyName(string name)
        {
            List<string> parts = name.Split('.').ToList();
            parts.RemoveAt(parts.Count - 1);
            return string.Join(".", parts);
        }
        private string GetTemplates()
        {
            List<string> parts = new List<string>();
            DirectoryInfo templateFolder = new DirectoryInfo(@"X:\BIM\2_Шаблоны");
            foreach (DirectoryInfo folder in templateFolder.GetDirectories())
            {
                if (KPLN_Loader.Preferences.User.Department.Id == 1 && (folder.Name != "1_АР" && folder.Name != "0_Общие шаблоны")) { continue; }
                if (KPLN_Loader.Preferences.User.Department.Id == 2 && (folder.Name != "2_КР" && folder.Name != "0_Общие шаблоны")) { continue; }
                if (KPLN_Loader.Preferences.User.Department.Id == 3 && (folder.Name == "1_АР" || folder.Name == "2_КР" || folder.Name == "0_Общие шаблоны")) { continue; }
                foreach (FileInfo file in folder.GetFiles())
                {
                    if (IsCopy(file.Name) || file.Extension.ToLower() != ".rte") { continue; }
                    parts.Add(string.Format("{0} - {1}={2}", folder.Name.Split('_').Last(), OnlyName(file.Name), file.FullName));
                }
            }
            return string.Join(",", parts);
        }
        private void OverwriteIni(string revitVersion)
        {
            FileInfo iniLocation = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), string.Format(@"AppData\Roaming\Autodesk\Revit\Autodesk Revit {0}\Revit.ini", revitVersion)));
            if (!iniLocation.Exists) { return; }
            INIManager manager = new INIManager(iniLocation.FullName);
            manager.WritePrivateString("Selection", "AllowPressAndDrag", "0");
            manager.WritePrivateString("Selection", "AllowFaceSelection", "0");
            manager.WritePrivateString("Selection", "AllowUnderlaySelection", "1");
#if Revit2020
            manager.WritePrivateString("DirectoriesRUS", "DefaultTemplate", GetTemplates());
            manager.WritePrivateString("DirectoriesENU", "DefaultTemplate", GetTemplates());
#endif
        }
        private List<SQLProject> GetProjects()
        {
            List<SQLProject> projects = new List<SQLProject>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Projects", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            projects.Add(new SQLProject(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(3), rdr.GetString(4)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception) { }
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return projects;
        }
        private List<SQLDepartment> GetDepartments()
        {
            List<SQLDepartment> departments = new List<SQLDepartment>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM SubDepartments", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            departments.Add(new SQLDepartment(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception) { }
            {
                try { sql.Close(); }
                catch (Exception) { }
            }
            return departments;
        }
    }
}
