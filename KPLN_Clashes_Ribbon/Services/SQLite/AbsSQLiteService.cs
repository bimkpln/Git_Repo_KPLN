using Dapper;
using KPLN_Clashes_Ribbon.Core.Reports;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    /// <summary>Повторяется только откатившаяся операция; ожидание выполняется без открытой транзакции.</summary>
    internal static class SQLiteRetryPolicy
    {
        private static readonly int[] RetryDelays = { 200, 500, 1000 };

        internal static SQLiteConnection CreateConnection(string connectionString)
        {
            var builder = new SQLiteConnectionStringBuilder(connectionString)
            {
                DefaultTimeout = 1,
                FailIfMissing = true
            };
            return new SQLiteConnection(builder.ConnectionString);
        }

        internal static async Task<T> RunAsync<T>(Func<T> attempt)
        {
            for (int index = 0; ; index++)
            {
                try
                {
                    return await Task.Run(attempt).ConfigureAwait(false);
                }
                catch (SQLiteException ex) when (IsBusy(ex))
                {
                    if (index == RetryDelays.Length)
                        throw new TimeoutException("База данных занята другим пользователем. Четыре попытки записи не завершились. Повторите действие позже.", ex);

                    // attempt уже вышел из using: соединение закрыто, транзакция откатилась.
                    await Task.Delay(RetryDelays[index]).ConfigureAwait(false);
                }
            }
        }

        private static bool IsBusy(SQLiteException exception)
        {
            int primaryCode = (int)exception.ResultCode & 0xff;
            return primaryCode == (int)SQLiteErrorCode.Busy || primaryCode == (int)SQLiteErrorCode.Locked;
        }
    }

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
                StringBuilder sb = new StringBuilder();
                if (parameters != null)
                {
                    // Беру параметр имени из динамика, если он есть
                    var prop = parameters.GetType().GetProperty("Name");
                    if (prop != null)
                    {
                        var name = prop.GetValue(parameters)?.ToString();
                        sb.Append($"Ошибка в отчёте по имени \"{name}\": {ex.Message}");
                    }
                    else
                        sb.Append($"{ex.Message}");
                }

                Print(sb.ToString(), MessageType.Error);
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
