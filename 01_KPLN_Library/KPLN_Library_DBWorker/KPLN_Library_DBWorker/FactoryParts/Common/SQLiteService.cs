using KPLN_Library_DBWorker.Core;
using System.Data.SQLite;

namespace KPLN_Library_DBWorker.FactoryParts.Common
{
    /// <summary>
    /// Общий сервис работы с листами БД - SQLite
    /// </summary>
    public class SQLiteService : DbServiceAbstr<SQLiteConnection, SQLiteException>
    {
        private protected SQLiteService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        protected override SQLiteConnection CreateConnection(string connectionString) => new SQLiteConnection(connectionString);

        protected override bool IsRetryable(SQLiteException ex) => ex.ErrorCode == (int)SQLiteErrorCode.Busy;

        protected override string DbProviderName => "SQLite";
    }
}
