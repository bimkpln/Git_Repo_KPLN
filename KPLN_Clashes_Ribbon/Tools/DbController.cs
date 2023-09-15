using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Tools
{
    internal static class DbController
    {
        private static readonly string _dbConnection = @"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_NwcReports.db;Version=3;";

        public static void UpdateReportLastChange(int report, ClashesMainCollection.KPItemStatus mainStatus)
        {
            string name = KPLN_Loader.Application.CurrentRevitUser.SystemName;

            if (mainStatus == ClashesMainCollection.KPItemStatus.Delegated)
            {
                SetReportValue(report, "KPItemStatus", 2);
            }
            else if (mainStatus == ClashesMainCollection.KPItemStatus.Closed)
            {
                SetReportValue(report, "KPItemStatus", 1);
            }
            else
            {
                SetReportValue(report, "KPItemStatus", 0);
            }

            SetReportValue(report, "DateLast", DateTime.Now.ToString());
            SetReportValue(report, "UserLast", name);
        }

        public static void UpdateGroupLastChange(int group)
        {
            string name = KPLN_Loader.Application.CurrentRevitUser.SystemName;
            if (GetGroupValueInteger(group, "KPItemStatus") == -1)
            {
                SetGroupValue(group, "KPItemStatus", 0);
            }
            SetGroupValue(group, "DateLast", DateTime.Now.ToString());
            SetGroupValue(group, "UserLast", name);
        }

        public static void SetInstanceValue(string path, int id, string parameter, string value)
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path));
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE Reports SET {0}='{1}' WHERE ID={2}", parameter, value, id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }

        public static void SetInstanceValue(string path, int id, string parameter, int value)
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path));
            try
            {
                db.Open();

                // Добавление столбца для делегирования для старых отчетов, если его ранее не было
                if (parameter == "DEPARTMENT")
                {
                    try
                    {
                        SQLiteCommand cmd_check = new SQLiteCommand("ALTER TABLE Reports ADD COLUMN DEPARTMENT INTEGER", db);
                        cmd_check.ExecuteNonQuery();
                    }
                    catch (SQLiteException) { }
                }

                //Вставляем данные
                SQLiteCommand cmd_insert = new SQLiteCommand($"UPDATE Reports SET {parameter}={value} WHERE ID={id}", db);
                cmd_insert.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Print(ex.Message, MessageType.Error);
            }
            finally
            {
                db.Close();
            }
        }

        public static int GetGroupValueInteger(int id, string parameter)
        {
            int value = -1;
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                using (SQLiteCommand cmd1 = new SQLiteCommand(string.Format("SELECT {0} FROM ReportGroups WHERE Id={1}", parameter, id), db))
                {
                    using (SQLiteDataReader rdr = cmd1.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            value = rdr.GetInt32(0);
                        }
                    }
                }
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
            return value;
        }

        public static int GetReportValueInteger(int id, string parameter)
        {
            int value = -1;
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                using (SQLiteCommand cmd1 = new SQLiteCommand(string.Format("SELECT {0} FROM Reports WHERE Id={1}", parameter, id), db))
                {
                    using (SQLiteDataReader rdr = cmd1.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            value = rdr.GetInt32(0);
                        }
                    }
                }
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
            return value;
        }

        public static void SetGroupValue(int id, string parameter, string value)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE ReportGroups SET {0}='{1}' WHERE Id={2}", parameter, value, id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }

        public static void SetGroupValue(int id, string parameter, int value)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE ReportGroups SET {0}={1} WHERE Id={2}", parameter, value.ToString(), id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }

        public static void SetReportValue(int id, string parameter, int value)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE Reports SET {0}={1} WHERE Id={2}", parameter, value.ToString(), id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
            }
            finally
            {
                db.Close();
            }
        }

        public static void SetReportValue(int id, string parameter, string value)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE Reports SET {0}='{1}' WHERE Id={2}", parameter, value, id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }

        public static void RemoveGroup(ReportGroup group)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd = new SQLiteCommand(string.Format("DELETE FROM ReportGroups WHERE Id={0}", group.Id), db);
                cmd.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception e)
            {
                PrintError(e);
                db.Close();
            }
        }

        public static void RemoveReport(Report report)
        {
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                using (SQLiteCommand cmd1 = new SQLiteCommand(string.Format("SELECT PathToReportInstance FROM Reports WHERE Id={0}", report.Id), db))
                {
                    using (SQLiteDataReader rdr = cmd1.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            FileInfo DbFile = new FileInfo(rdr.GetString(0));
                            if (DbFile.Exists)
                            {
                                try
                                {
                                    DbFile.Delete();
                                }
                                catch (Exception) { }
                            }

                        }
                    }
                }
                SQLiteCommand cmd = new SQLiteCommand(string.Format("DELETE FROM Reports WHERE Id={0}", report.Id), db);
                cmd.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception e)
            {
                PrintError(e);
                db.Close();
            }
        }

        public static ObservableCollection<ReportItemComment> GetComments(FileInfo path, ReportItem instance)
        {
            string value = string.Empty;
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path.FullName));
            try
            {
                db.Open();
                using (SQLiteCommand cmd1 = new SQLiteCommand(string.Format("SELECT COMMENTS FROM Reports WHERE Id={0}", instance.Id.ToString()), db))
                {
                    using (SQLiteDataReader rdr = cmd1.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            value = rdr.GetString(0);
                        }
                    }
                }
            }
            catch (Exception) { }
            db.Close();
            return ReportItemComment.ParseComments(value, instance);
        }

        public static void AddComment(string message, FileInfo path, ReportItem instance, int type)
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path.FullName));
            db.Open();
            List<string> value_parts = new List<string>();
            try
            {
                value_parts.Add(new ReportItemComment(message, type).ToString());
                foreach (ReportItemComment comment in instance.Comments)
                {
                    value_parts.Add(comment.ToString());
                }
                string value = string.Join(ClashesMainCollection.StringSeparatorItem, value_parts);
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE Reports SET COMMENTS='{0}' WHERE ID={1}", value, instance.Id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
            }
            catch (Exception) { }
            db.Close();
        }

        public static void RemoveComment(ReportItemComment comment_to_remove, FileInfo path, ReportItem instance)
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path.FullName));
            db.Open();
            List<string> value_parts = new List<string>();
            try
            {
                foreach (ReportItemComment comment in instance.Comments)
                {
                    if (comment.Message != comment_to_remove.Message || comment.Time != comment_to_remove.Time || comment.UserFullName != comment_to_remove.UserFullName)
                    {
                        value_parts.Add(comment.ToString());
                    }
                }
                string value = string.Join(ClashesMainCollection.StringSeparatorItem, value_parts);
                SQLiteCommand cmd_insert = new SQLiteCommand(string.Format("UPDATE Reports SET COMMENTS='{0}' WHERE ID={1}", value, instance.Id.ToString()), db);
                cmd_insert.ExecuteNonQuery();
            }
            catch (Exception) { }
            db.Close();
        }

        public static void AddGroup(string name, DBProject project)
        {
            string user = KPLN_Loader.Application.CurrentRevitUser.SystemName;
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(
                    "INSERT INTO ReportGroups (" +
                        "[ProjectId], " +
                        "[Name], " +
                        "[KPItemStatus], " +
                        "[DateCreated], " +
                        "[UserCreated], " +
                        "[DateLast], " +
                        "[UserLast]) " +
                    "VALUES(" +
                        "@ProjectId, " +
                        "@Name, " +
                        "@KPItemStatus, " +
                        "@DateCreated, " +
                        "@UserCreated, " +
                        "@DateLast, " +
                        "@UserLast)",
                    db);

                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "ProjectId", Value = project.Id });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "Name", Value = name });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KPItemStatus", Value = -1 });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "DateCreated", Value = DateTime.Now.ToString() });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "UserCreated", Value = user });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "DateLast", Value = DateTime.Now.ToString() });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "UserLast", Value = user });
                cmd_insert.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception)
            {
                db.Close();
            }
        }

        /// <summary>
        /// Генерация пути для хранения файла БД
        /// </summary>
        public static FileInfo GenerateNewPath()
        {
            int step = 0;
            FileInfo file;
            while (new FileInfo(Path.Combine(@"Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader\DB\NavisWorksReports", string.Format("nwc_report_{0}.db", step.ToString()))).Exists)
            {
                step++;
            }
            file = new FileInfo(Path.Combine(@"Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader\DB\NavisWorksReports", string.Format("nwc_report_{0}.db", step.ToString())));
            return file;
        }

        /// <summary>
        /// Создание БД для хранения отчетов
        /// </summary>
        public static void CreateDbFile(FileInfo file)
        {
            try
            {
                Thread thread = new Thread(() =>
                {
                    bool exist = File.Exists(string.Format(@"Data Source={0};Version=3;", file.FullName));
                    SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", file.FullName));
                    try
                    {
                        db.Open();
                        SQLiteCommand cmd_create = new SQLiteCommand(
                            "CREATE TABLE Reports (" +
                                "ID INTEGER PRIMARY KEY, " +
                                "NAME TEXT, " +
                                "IMAGE BLOB, " +
                                "IMAGE_PREVIEW BLOB, " +
                                "ELEMENT01 TEXT, " +
                                "ELEMENT02 TEXT, " +
                                "POINT TEXT, " +
                                "STATUS INTEGER, " +
                                "COMMENTS TEXT, " +
                                "GROUPID INTEGER, " +
                                "DEPARTMENT INTEGER)",
                            db);

                        cmd_create.ExecuteNonQuery();
                        db.Close();
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                        db.Close();
                    }
                })
                {
                    IsBackground = true,
                    Name = "sql_insert_data"
                };
                thread.Start();
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Заполнить отчет данными
        /// </summary>
        /// <param name="path">Путь к отчету</param>
        /// <param name="reports">Коллекция экземпляров класса-отчета</param>
        public static void FillReports(FileInfo path, ObservableCollection<ReportItem> reports)
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path.FullName));
            db.Open();
            try
            {
                foreach (ReportItem report in reports)
                {
                    SQLiteCommand cmd_insert = new SQLiteCommand(
                        "INSERT INTO Reports (" +
                            "[ID], " +
                            "[NAME], " +
                            "[IMAGE], " +
                            "[IMAGE_PREVIEW], " +
                            "[ELEMENT01], " +
                            "[ELEMENT02], " +
                            "[POINT], " +
                            "[STATUS], " +
                            "[COMMENTS], " +
                            "[GROUPID]) " +
                        "VALUES(" +
                            "@ID, " +
                            "@NAME, " +
                            "@IMAGE, " +
                            "@IMAGE_PREVIEW, " +
                            "@ELEMENT01, " +
                            "@ELEMENT02, " +
                            "@POINT, " +
                            "@STATUS, " +
                            "@COMMENTS, " +
                            "@GROUPID)",
                        db);

                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "ID", Value = report.Id });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "NAME", Value = report.Name });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "IMAGE", Value = report.Image, DbType = System.Data.DbType.Binary });
                    //cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "IMAGE_PREVIEW", Value = report.ImageData_Preview, DbType = System.Data.DbType.Binary });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "ELEMENT01", Value = string.Join("|", new string[] { report.Element_1_Id.ToString(), report.Element_1_Info }) });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "ELEMENT02", Value = string.Join("|", new string[] { report.Element_2_Id.ToString(), report.Element_2_Info }) });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "POINT", Value = report.Point });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "STATUS", Value = -1 });
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "COMMENTS", Value = ReportItem.GetCommentsString(report.Comments) }); ;
                    cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "GROUPID", Value = report.ParentGroupId }); ;
                    cmd_insert.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            { PrintError(e); }
            db.Close();
        }

        /// <summary>
        /// Генерация и заполнение файла БД под отчет
        /// </summary>
        /// <param name="name">Имя группы отчета</param>
        /// <param name="group">Экземпляр класса-группы</param>
        /// <param name="reports">Коллекция экземпляров класса-отчета</param>
        public static void AddReport(string name, ReportGroup group, ObservableCollection<ReportItem> reports)
        {
            string user = KPLN_Loader.Application.CurrentRevitUser.SystemName;
            SQLiteConnection db = new SQLiteConnection(_dbConnection);
            try
            {
                FileInfo path = GenerateNewPath();

                CreateDbFile(path);

                db.Open();
                SQLiteCommand cmd_insert = new SQLiteCommand(
                    "INSERT INTO Reports (" +
                        "[ReportGroupId], " +
                        "[Name], " +
                        "[KPItemStatus], " +
                        "[PathToReportInstance], " +
                        "[DateCreated], " +
                        "[UserCreated], " +
                        "[DateLast], " +
                        "[UserLast]) " +
                    "VALUES(" +
                        "@ReportGroupId, " +
                        "@Name, " +
                        "@KPItemStatus, " +
                        "@PathToReportInstance, " +
                        "@DateCreated, " +
                        "@UserCreated, " +
                        "@DateLast, " +
                        "@UserLast)",
                    db);

                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "ReportGroupId", Value = group.Id });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "Name", Value = name });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "KPItemStatus", Value = -1 });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "PathToReportInstance", Value = path.FullName });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "DateCreated", Value = DateTime.Now.ToString() });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "UserCreated", Value = user });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "DateLast", Value = DateTime.Now.ToString() });
                cmd_insert.Parameters.Add(new SQLiteParameter() { ParameterName = "UserLast", Value = user });
                cmd_insert.ExecuteNonQuery();
                db.Close();

                FillReports(path, reports);
            }
            catch (Exception e)
            {
                PrintError(e);
                db.Close();
            }
        }
    }
}
