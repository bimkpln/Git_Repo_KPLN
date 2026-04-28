using KPLN_Loader.Services.Abstract;
using MySqlConnector;
using NLog;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для работы с БД-онлайн
    /// </summary>
    internal sealed class MySQLService : DBServiceAbstr<MySqlConnection, MySqlException>
    {
        internal MySQLService(Logger logger, string dbPath) : base(logger, dbPath)
        {
            try
            {
                using (var conn = CreateConnection(dbPath))
                    conn.Open();
            }
            catch
            {
                throw;
            }
        }

        protected override MySqlConnection CreateConnection(string connectionStringOrPath) => new MySqlConnection(connectionStringOrPath);

        protected override bool IsRetryable(MySqlException ex) => true;

        protected override string DbProviderName => "MySQL";
    }
}