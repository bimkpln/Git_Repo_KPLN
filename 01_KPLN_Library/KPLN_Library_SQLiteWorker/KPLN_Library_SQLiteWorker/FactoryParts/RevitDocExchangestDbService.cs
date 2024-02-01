using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом RevitDocExchanges в БД
    /// </summary>
    public class RevitDocExchangestDbService : DbService
    {
        internal RevitDocExchangestDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        /// <summary>
        /// Получить коллекцию ВСЕХ активных файлов для обмена
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitActiveDocExchanges() =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} WHERE {nameof(DBRevitDocExchanges.IsActive)}='True';");

        /// <summary>
        /// Обновить статус IsActive документа по статусу проекта
        /// </summary>
        public void UpdateDBRevitDocExchanges_IsClosedByProject(DBProject dbProject) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                  $"SET {nameof(DBRevitDocExchanges.IsActive)}='{dbProject.IsClosed}' WHERE {nameof(DBRevitDocExchanges.ProjectId)}='{dbProject.Id}';");
    }
}
