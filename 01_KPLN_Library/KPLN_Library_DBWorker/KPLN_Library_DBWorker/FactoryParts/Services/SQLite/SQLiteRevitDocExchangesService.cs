using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using Dapper;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом RevitDocExchanges в БД
    /// </summary>
    public class SQLiteRevitDocExchangesService : SQLiteService
    {
        internal SQLiteRevitDocExchangesService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Создание новой сущности в БД
        /// </summary>
        /// <param name="docExchanges"></param>
        public int CreateDBRevitDocExchanges(DBRevitDocExchanges docExchanges) =>
            SaveConfiguration(docExchanges, true);
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
            SaveConfiguration(currentDocExc, false);
        #endregion

        /// <summary>
        /// Сохранение общих настроек экспорта. Настройки автозапуска хранятся в ModuleAutostart.
        /// </summary>
        private int SaveConfiguration(DBRevitDocExchanges configuration, bool create)
        {
            using (var connection = CreateConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction(System.Data.IsolationLevel.Serializable))
                {
                    int id = configuration.Id;
                    if (create)
                    {
                        id = connection.ExecuteScalar<int>(
                            $"INSERT INTO {_dbTableName} " +
                            "(ProjectId, RevitDocExchangeType, SettingName, SettingResultPath, SettingCountItem, SettingDBFilePath) " +
                            "VALUES (@ProjectId, @RevitDocExchangeType, @SettingName, @SettingResultPath, @SettingCountItem, @SettingDBFilePath) RETURNING Id;",
                            configuration, transaction);
                    }
                    else
                    {
                        connection.Execute(
                            $"UPDATE {_dbTableName} SET SettingName = @SettingName, SettingResultPath = @SettingResultPath, " +
                            "SettingCountItem = @SettingCountItem WHERE Id = @Id;",
                            configuration, transaction);
                    }

                    transaction.Commit();
                    return id;
                }
            }
        }

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
