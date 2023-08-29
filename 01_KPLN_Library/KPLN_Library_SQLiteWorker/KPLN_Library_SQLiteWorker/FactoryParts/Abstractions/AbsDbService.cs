using Dapper;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace KPLN_Library_SQLiteWorker.FactoryParts.Abstractions
{
    /// <summary>
    /// Абстрактный сервис работы с БД
    /// </summary>
    public abstract class AbsDbService
    {
        private protected string _connectionString;

        private protected AbsDbService(string connectionString)
        {
            _connectionString = connectionString;
        }

        private protected void ExecuteNonQuery(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                connection.Execute(query, parameters);
            }
        }

        private protected IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                return connection.Query<T>(query, parameters);
            }
        }
    }
}
