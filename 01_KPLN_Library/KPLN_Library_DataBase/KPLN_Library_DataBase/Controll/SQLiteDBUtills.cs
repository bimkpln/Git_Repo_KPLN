using KPLN_Library_DataBase.Collections;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Data.SQLite;

namespace KPLN_Library_DataBase.Controll
{
    public static class SQLiteDBUtills
    {
        private static string _sqlConnection = DbControll.MainDBConnection;

        public static ObservableCollection<DbUserInfo> GetUserInfo(ObservableCollection<DbDepartment> departments)
        {
            ObservableCollection<DbUserInfo> users = new ObservableCollection<DbUserInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            
            sql.ConnectionString = _sqlConnection;
            sql.Open();
            try
            {
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
            }
            catch (Exception ex)
            {
                sql.Close();
                throw new Exception($"{typeof(SQLiteDBUtills).FullName} get exception: {ex}\n");
            }
            sql.Close();

            return users;
        }
        
        public static ObservableCollection<DbDepartmentInfo> GetDepartmentInfo()
        {
            ObservableCollection<DbDepartmentInfo> departments = new ObservableCollection<DbDepartmentInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            
            sql.ConnectionString = _sqlConnection;
            sql.Open();
            try
            {
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
            }
            catch (Exception ex)
            {
                sql.Close();
                throw new Exception($"{typeof(SQLiteDBUtills).FullName} get exception: {ex}\n");
            }
            sql.Close();

            return departments;
        }
        
        public static ObservableCollection<DbSubDepartmentInfo> GetSubDepartmentInfo()
        {
            ObservableCollection<DbSubDepartmentInfo> subDepartments = new ObservableCollection<DbSubDepartmentInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            
            sql.ConnectionString = _sqlConnection;
            sql.Open();
            try
            {
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
            }
            catch (Exception ex)
            {
                sql.Close();
                throw new Exception($"{typeof(SQLiteDBUtills).FullName} get exception: {ex}\n");
            }
            sql.Close();

            return subDepartments;
        }
        
        public static ObservableCollection<DbProjectInfo> GetProjectInfo(ObservableCollection<DbUser> users)
        {
            ObservableCollection<DbProjectInfo> projects = new ObservableCollection<DbProjectInfo>();
            SQLiteConnection sql = new SQLiteConnection();

            sql.ConnectionString = _sqlConnection;
            sql.Open();
            try
            {
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
            }
            catch (Exception ex)
            {
                sql.Close();
                throw new Exception($"{typeof(SQLiteDBUtills).FullName} get exception: {ex}\n");
            }
            sql.Close();

            return projects;
        }
        
        public static ObservableCollection<DbDocumentInfo> GetDocumentInfo(ObservableCollection<DbSubDepartment> subdepartments, ObservableCollection<DbProject> projects)
        {
            ObservableCollection<DbDocumentInfo> documents = new ObservableCollection<DbDocumentInfo>();
            SQLiteConnection sql = new SQLiteConnection();
            
            sql.ConnectionString = _sqlConnection;
            sql.Open();
            try
            {
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
            }
            catch (Exception ex) 
            { 
                sql.Close();
                throw new Exception($"{typeof(SQLiteDBUtills).FullName} get exception: {ex}\n"); 
            }
            sql.Close();

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
