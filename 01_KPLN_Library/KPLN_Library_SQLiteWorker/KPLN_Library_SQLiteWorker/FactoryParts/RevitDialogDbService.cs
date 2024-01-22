using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом RevitDialog в БД
    /// </summary>
    public class RevitDialogDbService : DbService
    {
        internal RevitDialogDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
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
    }
}
