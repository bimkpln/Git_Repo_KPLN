﻿using Autodesk.Revit.UI;
using Dapper;
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
        private protected string _dbTableName;

        /// <param name="connectionString">Строка подключения</param>
        /// <param name="dbTableName">Имя таблицы в БД</param>
        private protected DbService(string connectionString, string dbTableName)
        {
            _connectionString = connectionString;
            _dbTableName = dbTableName;
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
