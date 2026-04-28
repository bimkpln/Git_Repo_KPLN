using KPLN_Loader.Services.Abstract;
using NLog;
using System.Data.SQLite;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для работы с БД - SQLite
    /// </summary>
    internal sealed class SQLiteService : DBServiceAbstr<SQLiteConnection, SQLiteException>
    {
        internal SQLiteService(Logger logger, string dbPath) : base(logger, "Data Source=" + dbPath + ";Version=3;")
        {
        }

        protected override SQLiteConnection CreateConnection(string connectionStringOrPath) => new SQLiteConnection(connectionStringOrPath);

        protected override bool IsRetryable(SQLiteException ex) => ex.ErrorCode == (int)SQLiteErrorCode.Busy;

        protected override string DbProviderName => "SQLite";
    }
}
