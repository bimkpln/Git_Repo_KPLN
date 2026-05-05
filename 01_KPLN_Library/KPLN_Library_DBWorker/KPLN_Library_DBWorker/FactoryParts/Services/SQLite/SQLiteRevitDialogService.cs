using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом RevitDialog в БД
    /// </summary>
    public class SQLiteRevitDialogService : SQLiteService
    {
        internal SQLiteRevitDialogService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        /// <summary>
        /// Получить все RevitDialog
        /// </summary>
        public IEnumerable<DBRevitDialog> GetDBRevitDialogs() =>
            ExecuteQuery<DBRevitDialog>(
                 $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить RevitDialog по DialogId
        /// </summary>
        public IEnumerable<DBRevitDialog> GetDBRevitDialog_ByDialogId(string dialogId) =>
            ExecuteQuery<DBRevitDialog>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDialog.DialogId)}='{dialogId}';");

        /// <summary>
        /// Получить RevitDialog по полному совпадению Message
        /// </summary>
        public IEnumerable<DBRevitDialog> GetDBRevitDialog_ByEqualMessage(string message) =>
            ExecuteQuery<DBRevitDialog>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDialog.Message)}='{message}';");
    }
}
