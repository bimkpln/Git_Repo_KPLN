using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using KPLN_Loader.Core.SQLiteData;
using System.Collections.Generic;
using System.Linq;

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

        #region Create
        /// <summary>
        /// Создание новой сущности в БД
        /// </summary>
        /// <param name="docExchanges"></param>
        public void CreateDBRevitDocExchanges(DBRevitDocExchanges docExchanges) =>
            ExecuteNonQuery($"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBRevitDocExchanges.ProjectId)}, {nameof(DBRevitDocExchanges.RevitDocExchangeType)}, {nameof(DBRevitDocExchanges.SettingName)}, {nameof(DBRevitDocExchanges.SettingDBFilePath)}) " +
                    $"VALUES (@{nameof(DBRevitDocExchanges.ProjectId)}, @{nameof(DBRevitDocExchanges.RevitDocExchangeType)}, @{nameof(DBRevitDocExchanges.SettingName)}, @{nameof(DBRevitDocExchanges.SettingDBFilePath)});",
                docExchanges);
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ активных файлов-конфигураций для обмена
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitActiveDocExchanges() =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} WHERE {nameof(DBRevitDocExchanges.IsActive)}='True';");

        /// <summary>
        /// Получить конфигурации по типу обмена и по проекту
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitDocExchanges_ByExchangeTypeANDDBProject(RevitDocExchangeEnum revitDocExchangeEnum, DBProject dbProject) =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.IsActive)}='True' " +
                $"AND {nameof(DBRevitDocExchanges.RevitDocExchangeType)}='{revitDocExchangeEnum}' " +
                $"AND {nameof(DBRevitDocExchanges.ProjectId)}='{dbProject.Id}';");
        #endregion

        #region Update
        /// <summary>
        /// Обновить статус IsActive документа по статусу проекта
        /// </summary>
        public void UpdateDBRevitDocExchanges_IsClosedByProject(DBProject dbProject) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                  $"SET {nameof(DBRevitDocExchanges.IsActive)}='{dbProject.IsClosed}' WHERE {nameof(DBRevitDocExchanges.ProjectId)}='{dbProject.Id}';");
        #endregion

        #region Delete
        /// <summary>
        /// Удалить файл-конфигурацию для обмена
        /// </summary>
        public void DeleteDBRevitDocExchange_ById(int id) =>
            ExecuteNonQuery($"DELETE FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)} = {id}");
        #endregion

    }
}
