using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом RevitDocExchanges в БД
    /// </summary>
    public class RevitDocExchangesDbService : DbService
    {
        internal RevitDocExchangesDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Создание новой сущности в БД
        /// </summary>
        /// <param name="docExchanges"></param>
        public int CreateDBRevitDocExchanges(DBRevitDocExchanges docExchanges) =>
            ExecuteQuery<int>(
                $"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBRevitDocExchanges.ProjectId)}, {nameof(DBRevitDocExchanges.RevitDocExchangeType)}, {nameof(DBRevitDocExchanges.SettingName)}, {nameof(DBRevitDocExchanges.SettingResultPath)}, {nameof(DBRevitDocExchanges.SettingCountItem)}, {nameof(DBRevitDocExchanges.SettingDBFilePath)}) " +
                    $"VALUES (@{nameof(DBRevitDocExchanges.ProjectId)}, @{nameof(DBRevitDocExchanges.RevitDocExchangeType)}, @{nameof(DBRevitDocExchanges.SettingName)}, @{nameof(DBRevitDocExchanges.SettingResultPath)}, @{nameof(DBRevitDocExchanges.SettingCountItem)}, @{nameof(DBRevitDocExchanges.SettingDBFilePath)})" +
                    $"RETURNING Id;",
                docExchanges)
            .FirstOrDefault();
        #endregion

        #region Read
        /// <summary>
        /// Получить конфигурацию по id
        /// </summary>
        public DBRevitDocExchanges GetDBRevitDocExchanges_ById(int id) =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)}='{id}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить конфигурации по КОЛЛЕКЦИИ id
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitDocExchanges_ByIdCol(IEnumerable<int> ids) =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)} IN @Ids;",
                new { Ids = ids });

        /// <summary>
        /// Получить коллекцию ВСЕХ активных файлов-конфигураций для обмена
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitActiveDocExchanges() =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить конфигурации по типу обмена и по проекту
        /// </summary>
        public IEnumerable<DBRevitDocExchanges> GetDBRevitDocExchanges_ByExchangeTypeANDDBProject(RevitDocExchangeEnum revitDocExchangeEnum, DBProject dbProject) =>
            ExecuteQuery<DBRevitDocExchanges>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.RevitDocExchangeType)}='{revitDocExchangeEnum}' " +
                $"AND {nameof(DBRevitDocExchanges.ProjectId)}='{dbProject.Id}';");
        #endregion

        #region Update
        /// <summary>
        /// Обновить сущность по значениям из указанной
        /// </summary>
        /// <param name="currentDocExc"></param>
        public void UpdateDBRevitDocExchanges_ByDBRevitDocExchange(DBRevitDocExchanges currentDocExc) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                $"SET {nameof(DBRevitDocExchanges.SettingName)}='{currentDocExc.SettingName}', " +
                $"{nameof(DBRevitDocExchanges.SettingResultPath)}='{currentDocExc.SettingResultPath}', " +
                $"{nameof(DBRevitDocExchanges.SettingCountItem)}='{currentDocExc.SettingCountItem}' " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)}='{currentDocExc.Id}';");
        #endregion

        #region Delete
        /// <summary>
        /// Удалить файл-конфигурацию для обмена по id
        /// </summary>
        public void DeleteDBRevitDocExchange_ById(int id) =>
            ExecuteNonQuery($"DELETE FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)} = {id}");

        /// <summary>
        /// Удалить файл-конфигурацию для обмена по КОЛЛЕКЦИИ id
        /// </summary>
        public void DeleteDBRevitDocExchange_ByIdColl(IEnumerable<int> ids) =>
            ExecuteNonQuery($"DELETE FROM {_dbTableName} " +
                $"WHERE {nameof(DBRevitDocExchanges.Id)} IN @Ids;",
                new { Ids = ids });
        #endregion
    }
}
