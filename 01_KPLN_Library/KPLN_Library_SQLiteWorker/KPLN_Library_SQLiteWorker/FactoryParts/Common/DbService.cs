using Autodesk.Revit.UI;
using Dapper;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace KPLN_Library_SQLiteWorker.FactoryParts.Common
{
    /// <summary>
    /// Общий сервис работы с листами БД
    /// </summary>
    public class DbService
    {
        private protected string _connectionString;
        private protected DB_Enumerator _dBEnumerator;
        private protected string _dbTableName;

        /// <param name="connectionString">Строка подключения</param>
        /// <param name="dbEnumerator">Таблица в БД</param>
        private protected DbService(string connectionString, DB_Enumerator dbEnumerator)
        {
            _connectionString = connectionString;
            _dBEnumerator = dbEnumerator;
            _dbTableName = _dBEnumerator.ToString();
        }

        private protected void ExecuteNonQuery(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    connection.Execute(query, parameters);
                }
            }
            catch (Exception ex) 
            {
                TaskDialog errorTd = new TaskDialog("[KPLN]: Ошибка работы с БД")
                {
                    MainContent = ex.Message,
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                errorTd.Show();
            }
        }

        private protected IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            try
            {
                using (IDbConnection connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();
                    return connection.Query<T>(query, parameters);
                }
            }
            catch (Exception ex)
            {
                TaskDialog errorTd = new TaskDialog("[KPLN]: Ошибка работы с БД")
                {
                    MainContent = ex.Message,
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    CommonButtons = TaskDialogCommonButtons.Ok,
                };
                errorTd.Show();

                return null;
            }
        }
    }
}
