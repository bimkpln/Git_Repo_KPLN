using KPLN_Loader.Common;
using static KPLN_Loader.Preferences;
using static KPLN_Loader.Output.Output;

using System;
using System.Collections.Generic;
using System.Data.SQLite;

namespace KPLN_Loader
{
    public class Tools_SQL
    {
        private SQLiteConnection sql = new SQLiteConnection();
        public Tools_SQL()
        {
            sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
        }
        private string GetCurentTime()
        {
            DateTime time = DateTime.Now;
            string[] parts = new string[] { time.Year.ToString(), time.Month.ToString(), time.Day.ToString(),
            time.Hour.ToString(), time.Minute.ToString(), time.Second.ToString() };
            return string.Join(":", parts);
        }
        private bool ProjectIdInList(SQLProjectInfo project, List<SQLProjectInfo> projectList)
        {
            foreach (SQLProjectInfo p in projectList)
            {
                if (p.Id == project.Id) { return true; }
            }
            return false;
        }
        public List<SQLDepartmentInfo> GetDepartments()
        {
            List<SQLDepartmentInfo> departments = new List<SQLDepartmentInfo>();
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Departments", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetInt32(0) == -1) { continue; }
                            departments.Add(new SQLDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            try
            {
                sql.Close();
            }
            catch (Exception) { }
            return departments;

        }
        public void GetUserProjects(string systemName, bool initialization = false)
        {
            string log = "";
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT * FROM Projects WHERE Users LIKE '%{0}%'", systemName), sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            SQLProjectInfo newProject = new SQLProjectInfo(rdr.GetInt32(0), rdr.GetString(1));
                            if (!ProjectIdInList(newProject, User_Projects))
                            {
                                User_Projects.Add(newProject);
                                log += string.Format("#{0} ", newProject.Name);
                                if (!initialization) { Print(string.Format("Вас добавили в проект #{0}", rdr.GetString(1)), MessageType.System_OK); }
                            }
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception e) { PrintError(e); }
            if (initialization)
            {
                Print(log, MessageType.System_Regular);
            }
            try { sql.Close(); } catch (Exception) { }
        }
        public void GetUsers(int loop = 1)
        {
            Preferences.Users.Clear();
            try
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
                sql.Open();
                List<SQLDepartmentInfo> departments = new List<SQLDepartmentInfo>();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Departments", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            departments.Add(new SQLDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2)));
                        }
                    }
                }
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Users", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            SQLUserInfo user = new SQLUserInfo(rdr.GetString(0), rdr.GetString(1), rdr.GetString(3), rdr.GetString(2), rdr.GetString(5));
                            int department_id = rdr.GetInt32(4);
                            foreach (SQLDepartmentInfo dep in departments)
                            {
                                if (dep.Id == department_id) { user.Department = dep; }
                            }
                            Users.Add(user);
                        }
                    }
                }

                sql.Close();
            }
            catch (Exception)
            {
                if (loop < 10)
                {
                    GetUsers(loop++);
                }
            }
        }
        public SQLUserInfo GetUser(string systemName, int loop = 1)
        {
            SQLUserInfo User = null;
            try
            {
                int department_id = -1;
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Users", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetString(0) == systemName)
                            {
                                User = new SQLUserInfo(systemName, rdr.GetString(1), rdr.GetString(3), rdr.GetString(2), rdr.GetString(5));
                                department_id = rdr.GetInt32(4);
                            }
                        }
                    }
                }
                using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT * FROM Departments WHERE Id = {0}", department_id.ToString()), sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            User.Department = new SQLDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                if (loop < 10)
                {
                    return GetUser(systemName, loop++);
                }
            }
            try { sql.Close(); } catch (Exception) { }
            return User;
        }
        public void GetUserData(string systemName, int loop=1)
        {
            try
            {
                int department_id = -1;
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Users", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetString(0) == systemName)
                            {
                                User = new SQLUserInfo(systemName, rdr.GetString(1), rdr.GetString(3), rdr.GetString(2), rdr.GetString(5));
                                department_id = rdr.GetInt32(4);
                            }
                        }
                    }
                }
                using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT * FROM Departments WHERE Id = {0}", department_id.ToString()), sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            User.Department = new SQLDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2));
                        }
                    }
                }
                SQLiteCommand cmd_update = new SQLiteCommand(sql);
                cmd_update.CommandText = string.Format("UPDATE Users SET Connection = '{0}' WHERE SystemName = '{1}'", GetCurentTime(), systemName);
                cmd_update.ExecuteNonQuery();
                sql.Close();
            }
            catch (Exception)
            {
                if (loop < 10)
                {
                    GetUserData(systemName, loop++);
                }
            }
            try { sql.Close(); } catch (Exception) { }
        }
        public bool IfUserExist(string username, int loop = 1)
        {
            if (loop > 1) { Print(string.Format("Попытка найти текущего пользователя {0}...", loop.ToString()), MessageType.Regular); }

            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT SystemName FROM Users", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            if (rdr.GetString(0) == username)
                            { return true; }
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception e)
            {
                if (loop < 10)
                {
                    return IfUserExist(username, loop++);
                }
                else
                {
                    Print(string.Format("Не удалось проверить наличие пользователя : {0}", e.Message), MessageType.Error);
                }
            }
            try { sql.Close(); } catch (Exception) { }
            return false;
        }
        public void CreateUser(string systemName, string name, string family, string surname, int department, int loop = 1)
        {
            if (loop > 1) { Print(string.Format("Попытка создать нового пользователя {0}...", loop.ToString()), MessageType.Regular); }
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = sql.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO Users ([SystemName], [Name], [Family], [Surname], [Department], [Status], [Connection], [Sex]) VALUES (@SystemName, @Name, @Family, @Surname, @Department, @Status, @Connection, @Sex)";
                    cmd.Parameters.Add(new SQLiteParameter("@SystemName") { Value = systemName });
                    cmd.Parameters.Add(new SQLiteParameter("@Name") { Value = name });
                    cmd.Parameters.Add(new SQLiteParameter("@Family") { Value = family });
                    cmd.Parameters.Add(new SQLiteParameter("@Surname") { Value = surname });
                    cmd.Parameters.Add(new SQLiteParameter("@Department") { Value = department });
                    cmd.Parameters.Add(new SQLiteParameter("@Status") { Value = ("New user") });
                    cmd.Parameters.Add(new SQLiteParameter("@Connection") { Value = GetCurentTime() });
                    cmd.Parameters.Add(new SQLiteParameter("@Sex") { Value = "Undefined" });
                    cmd.ExecuteNonQuery();
                }
                sql.Close();
            }
            catch (Exception e)
            {
                if (loop < 10)
                {
                    CreateUser(systemName, name, family, surname, department, loop++);
                }
                else
                {
                    Print(string.Format("Не удалось создать нового пользователя : {0}", e.Message), MessageType.Error);
                }
            }
            try { sql.Close(); } catch (Exception) { }
        }
        public void UpdateStatusMessage(int id, MessageDialogResult result)
        {
            string value = "Pending";
            switch (result)
            {
                case MessageDialogResult.Close:
                    value = "Close";
                    break;
                case MessageDialogResult.None:
                    value = "None";
                    break;
                case MessageDialogResult.Ok:
                    value = "Ok";
                    break;
                case MessageDialogResult.Read:
                    value = "Read";
                    break;
                case MessageDialogResult.Thx:
                    value = "Thx";
                    break;
                case MessageDialogResult.Pending:
                    value = "Pending";
                    break;
                case MessageDialogResult.Shown:
                    value = "Shown";
                    break;
            }
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                SQLiteCommand cmd = new SQLiteCommand(string.Format("UPDATE Users SET Status = '{1}' WHERE Id = '{0}'", id, value), sql);
                cmd.ExecuteNonQuery();
                sql.Close();
            }
            catch (Exception) { }
            try { sql.Close(); } catch (Exception) { }
        }
        public List<SQLModuleInfo> GetModules(string department, string table, string version, string projectId)
        {
            List<SQLModuleInfo> FoundedModules = new List<SQLModuleInfo>();
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT * FROM {0} WHERE (Department = {1} OR Department = -1) AND SupportedVersions LIKE '%{2}%' AND Enabled = 'True' AND Project = {3}", table, department, version, projectId), sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            try
                            {
                                FoundedModules.Add(new SQLModuleInfo(rdr.GetInt32(0), rdr.GetInt32(1), rdr.GetString(2), rdr.GetString(3), rdr.GetString(4)));
                            }
                            catch (Exception e) { PrintError(e); }
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception e)
            { PrintError(e); }
            try { sql.Close(); } catch (Exception) { }
            return FoundedModules;
        }
        public void UpdateUserConnection(string systemName, string table)
        {
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                SQLiteCommand cmd_update = new SQLiteCommand(sql);
                cmd_update.CommandText = string.Format("UPDATE {0} SET Connection = '{1}' WHERE SystemName = '{2}'", table, GetCurentTime(), systemName);
                cmd_update.ExecuteNonQuery();
                sql.Close();
            }
            catch (Exception)
            { }
            try { sql.Close(); } catch (Exception) { }
        }
        public void CreateLogMessage(string value)
        {
            string[] values = value.Split('*');
            try
            {
                try { sql.Close(); } catch (Exception) { }
                sql.Open();
                using (SQLiteCommand cmd = sql.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO Modules_Log ([User], [Name], [Date], [Log], [DocumentPath]) VALUES (@User, @Name, @Date, @Log, @DocumentPath)";
                    cmd.Parameters.Add(new SQLiteParameter("@User") { Value = values[0] });
                    cmd.Parameters.Add(new SQLiteParameter("@Name") { Value = values[1] });
                    cmd.Parameters.Add(new SQLiteParameter("@Date") { Value = values[2] });
                    cmd.Parameters.Add(new SQLiteParameter("@Log") { Value = values[3] });
                    cmd.Parameters.Add(new SQLiteParameter("@DocumentPath") { Value = values[4] });
                    cmd.ExecuteNonQuery();
                }
                sql.Close();
            }
            catch (Exception) { }
            try { sql.Close(); } catch (Exception) { }
        }
    }
}
