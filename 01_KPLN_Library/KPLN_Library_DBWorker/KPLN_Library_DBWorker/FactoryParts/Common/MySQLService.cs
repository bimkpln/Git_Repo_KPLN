using KPLN_Library_DBWorker.Core;
using MySqlConnector;

namespace KPLN_Library_DBWorker.FactoryParts.Common
{
    /// <summary>
    /// Общий сервис работы с листами БД - MySQL
    /// </summary>
    public class MySQLService : DbServiceAbstr<MySqlConnection, MySqlException>
    {
        internal MySQLService(string dbPath, DBEnumerator dbEnumerator) : base(dbPath, dbEnumerator)
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

        protected override string SqlTrueLiteral => "TRUE";

        protected override string SqlFalseLiteral => "FALSE";
    }
}
