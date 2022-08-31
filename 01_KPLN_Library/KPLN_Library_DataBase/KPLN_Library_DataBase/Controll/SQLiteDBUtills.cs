using KPLN_DataBase.Collections;
using System;
using System.Collections.ObjectModel;
using System.Data.SQLite;

namespace KPLN_DataBase.Controll
{
    public static class SQLiteDBUtills
    {
        public static ObservableCollection<DbUserInfo> GetUserInfo(ObservableCollection<DbDepartment> departments)
        {
            ObservableCollection<DbUserInfo> users = new ObservableCollection<DbUserInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Users", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            users.Add(new DbUserInfo(rdr.GetString(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3), rdr.GetInt32(4), rdr.GetString(5), rdr.GetString(6), rdr.GetString(7), departments));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return users;
        }
        public static ObservableCollection<DbDepartmentInfo> GetDepartmentInfo()
        {
            ObservableCollection<DbDepartmentInfo> departments = new ObservableCollection<DbDepartmentInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = string.Format(@"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader.db;Version=3;");
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Departments", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            departments.Add(new DbDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return departments;
        }
        public static ObservableCollection<DbSubDepartmentInfo> GetSubDepartmentInfo()
        {
            ObservableCollection<DbSubDepartmentInfo> subDepartments = new ObservableCollection<DbSubDepartmentInfo>();
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
                            subDepartments.Add(new DbSubDepartmentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3)));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return subDepartments;
        }
        public static ObservableCollection<DbProjectInfo> GetProjectInfo(ObservableCollection<DbUser> users)
        {
            ObservableCollection<DbProjectInfo> projects = new ObservableCollection<DbProjectInfo>();
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
                            projects.Add(new DbProjectInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3), rdr.GetString(4), users));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return projects;
        }
        public static ObservableCollection<DbDocumentInfo> GetDocumentInfo(ObservableCollection<DbSubDepartment> subdepartments, ObservableCollection<DbProject> projects)
        {
            ObservableCollection<DbDocumentInfo> documents = new ObservableCollection<DbDocumentInfo>();
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
                            documents.Add(new DbDocumentInfo(rdr.GetInt32(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(5), rdr.GetInt32(4), rdr.GetInt32(3), projects, subdepartments));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            return documents;
        }
        public static bool SetValue(DbElement element, string parameterName, string parameterValue)
        {
            try
            {
                return true;
            }
            catch (Exception)
            { }
            return false;
        }
        public static bool SetValue(DbElement element, string parameterName, int parameterValue)
        {
            try
            {
                return true;
            }
            catch (Exception)
            { }
            return false;
        }

    }
}
