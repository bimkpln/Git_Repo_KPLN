using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    public abstract class AbsSQLiteService
    {
        private protected string _dbConnectionPath;

        /// <summary>
        /// Текущее время в формате 'yyyy/MM/dd_HH:mm'
        /// </summary>
        private protected string CurrentTime
        {
            get => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
        }

        private protected void ExecuteNonQuery(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_dbConnectionPath))
                {
                    connection.Open();
                    connection.Execute(query, parameters);
                }
            }
            catch (Exception ex) { PrintError(ex); }
        }

        private protected int ExecuteInsertWithId(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_dbConnectionPath))
                {
                    connection.Open();
                    return connection.ExecuteScalar<int>(query, parameters);
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return -1;
            }
        }

        private protected void ExecuteNonQuery_Parameters(string query, IEnumerable<object> parameters)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_dbConnectionPath))
                {
                    connection.Open();
                    foreach (object parameter in parameters)
                    {
                        connection.Execute(query, parameter);
                    }
                }
            }
            catch (Exception ex) { PrintError(ex); }
        }

        private protected IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_dbConnectionPath))
                {
                    connection.Open();
                    return connection.Query<T>(query, parameters);
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return null;
            }
        }
    }
}
