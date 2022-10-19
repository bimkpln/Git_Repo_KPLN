using Autodesk.Revit.DB;
using KPLN_Scoper.Common;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using static KPLN_Scoper.Common.Collections;
using static KPLN_Loader.Output.Output;

namespace KPLN_Scoper.Tools
{
    public static class ActivityManager
    {
        public static ConcurrentQueue<ActivityInfo> ActivityBag = new ConcurrentQueue<ActivityInfo>();
        public static ConcurrentQueue<ActivityInfo> ActivityApprovedBag = new ConcurrentQueue<ActivityInfo>();
        public static int NullActions = 0;
        public static int ActivityActions = 0;
        
        public static Document ActiveDocument { get; set; }
        
        private static Timer _timer_small { get; set; }
        
        private static Timer _timer_big { get; set; }
        
        public static void Destroy()
        {
            _timer_small.Dispose();
            _timer_big.Dispose();
        }
        
        public static void Run()
        {
            var autoEvent_little = new AutoResetEvent(true);
            int time = 1000 * 60;
            _timer_small = new Timer(Update, autoEvent_little, time, time);
            var autoEvent_big = new AutoResetEvent(true);
            time = 1000 * 60 * 30;
            _timer_big = new Timer(Synchronize, autoEvent_big, time, time);
        }
        
        /// <summary>
        /// Запись всех данных за сессию (между синхронизациями) в БД
        /// </summary>
        public static void Synchronize(Object stateInfo)
        {
            Thread t = new Thread(() =>
            {
                try
                {
                    while (!ActivityApprovedBag.IsEmpty)
                    {
                        ActivityInfo dequeuesInfo;
                        if (ActivityApprovedBag.TryDequeue(out dequeuesInfo))
                        {
                            int prjId = dequeuesInfo.ProjectId;
                            int docId = dequeuesInfo.DocumentId;
                            string docTitle = dequeuesInfo.DocumentTitle;

                            AddValueToDb(KPLN_Loader.Preferences.User.SystemName, prjId, docId, dequeuesInfo.Time, dequeuesInfo.Type, docTitle, dequeuesInfo.Value);
                        }
                    }
                }
                catch (Exception) { }
            });
            t.IsBackground = true;
            t.Start();
        }
        
        /// <summary>
        /// Обновление данных по документу
        /// </summary>
        /// <param name="stateInfo"></param>
        public static void Update(Object stateInfo)
        {

            Thread t = new Thread(() =>
            {
                try
                {
                    ActivityInfo info = null;
                    if (ActiveDocument != null)
                    {
                        try
                        {
                            info = new ActivityInfo(ActiveDocument, BuiltInActivity.ActiveDocument);
                        }
                        catch (Exception)
                        {
                            info = null;
                        }
                    }

                    //Обработка активного документа
                    if (info != null)
                    {
                        ActivityApprovedBag.Enqueue(new ActivityInfo(info.DocumentId, info.ProjectId, BuiltInActivity.ActiveDocument, info.DocumentTitle, NullActions));
                        NullActions++;
                    }
                    
                    //Обработка активности с файлом
                    List<ActivityInfo> activitieInfos = new List<ActivityInfo>();
                    while (!ActivityBag.IsEmpty)
                    {
                        ActivityInfo dequeuesInfo;
                        if (ActivityBag.TryDequeue(out dequeuesInfo))
                        {
                            activitieInfos.Add(dequeuesInfo);
                        }
                    }
                    

                    //CHANGED DOCUMENTS
                    bool on_document_changed_found = false;
                    List<int> temp_doc = new List<int>();
                    List<int> temp_proj = new List<int>();
                    List<string> temp_titles = new List<string>();
                    ActivityActions = 0;
                    foreach (ActivityInfo i in activitieInfos)
                    {
                        if (i.Type == BuiltInActivity.DocumentChanged)
                        {
                            temp_doc.Add(i.DocumentId);
                            temp_proj.Add(i.ProjectId);
                            temp_titles.Add(i.DocumentTitle);
                            on_document_changed_found = true;
                            ActivityActions++;
                        }
                    }
                    if (on_document_changed_found)
                    {
                        int max_doc = MaxFrequent(temp_doc);
                        int max_proj = MaxFrequent(temp_proj);
                        string max_title = MaxFrequent(temp_titles);
                        NullActions = 0;
                        ActivityApprovedBag.Enqueue(new ActivityInfo(max_doc, max_proj, BuiltInActivity.DocumentChanged, max_title, ActivityActions));
                        ActivityActions = 0;
                        return;
                    }
                    
                    //SYNCHRONIZED DOCUMENTS
                    bool on_document_synchronized_found = false;
                    temp_doc = new List<int>();
                    temp_proj = new List<int>();
                    temp_titles = new List<string>();
                    foreach (ActivityInfo i in activitieInfos)
                    {
                        if (i.Type == BuiltInActivity.DocumentSynchronized)
                        {
                            temp_doc.Add(i.DocumentId);
                            temp_proj.Add(i.ProjectId);
                            temp_titles.Add(i.DocumentTitle);
                            on_document_synchronized_found = true;
                        }
                    }
                    if (on_document_synchronized_found)
                    {
                        int max_doc = MaxFrequent(temp_doc);
                        int max_proj = MaxFrequent(temp_proj);
                        string max_title = MaxFrequent(temp_titles);
                        NullActions = 0;
                        ActivityApprovedBag.Enqueue(new ActivityInfo(max_doc, max_proj, BuiltInActivity.DocumentSynchronized, max_title, NullActions));
                        return;
                    }
                }
                catch (Exception)
                { }

            });
            t.IsBackground = true;
            t.Start();
        }
        
        public static int CountValue<T>(List<T> list, T value)
        {
            int count = 0;
            foreach (T i in list)
            {
                if (i.Equals(value))
                {
                    count++;
                }
            }
            return count;
        }
        
        public static T MaxFrequent<T>(List<T> list)
        {
            T value = default(T);
            int max = -1;
            foreach (T i in list)
            {
                int current = CountValue(list, i);
                if (current > max)
                {
                    max = current;
                    value = i;
                }
            }
            return value;
        }
        
        private static FileInfo GetDbPath()
        {
            string base_name = string.Format("ADB_{0}_{1}_{2}.db", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString());
            string base_path = @"Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader\DB\ActivityData";
            string path = Path.Combine(base_path, base_name);
            FileInfo file = new FileInfo(path);
            if (file.Exists)
            {
                return file;
            }
            else
            {
                SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", file.FullName));
                try
                {
                    db.Open();
                    SQLiteCommand cmd_create = new SQLiteCommand("CREATE TABLE DATA (ID INTEGER PRIMARY KEY AUTOINCREMENT, USER TEXT, PROJECT INTEGER, DOCUMENT INTEGER, TIME TEXT, TYPE TEXT, SESSION TEXT, TITLE TEXT, VALUE REAL)", db);
                    cmd_create.ExecuteNonQuery();
                    db.Close();
                    return file;
                }
                catch (Exception)
                {
                    db.Close();
                    return null;
                }
            }
        }

        private static void AddValueToDb(string user, int project, int document, string time, BuiltInActivity type, string title, double value)
        {
            FileInfo dbFile = GetDbPath();
            if (dbFile == null) 
            {
                return; 
            }
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", dbFile.FullName));
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand("INSERT INTO DATA ([USER], [PROJECT], [DOCUMENT], [TIME], [TYPE], [SESSION], [TITLE], [VALUE]) VALUES(@USER, @PROJECT, @DOCUMENT, @TIME, @TYPE, @SESSION, @TITLE, @VALUE)", db);
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "USER", Value = user });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "PROJECT", Value = project });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "DOCUMENT", Value = document });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TIME", Value = time });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TYPE", Value = type.ToString("G") });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "SESSION", Value = ModuleData.SessionGuid.ToString() });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "TITLE", Value = title });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "VALUE", Value = Math.Round(value, 2) });
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }
    }
}
