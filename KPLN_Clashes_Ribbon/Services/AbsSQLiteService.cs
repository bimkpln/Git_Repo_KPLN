using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Services
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

        private protected IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_dbConnectionPath))
                {
                    connection.Open();
                    IEnumerable<T> data = connection.Query<T>(query, parameters);
                    if (data != null)
                        return data;

                    return null;
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
