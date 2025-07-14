using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом RevitDocExchanges в БД
    /// </summary>
    public class ModuleAutostartDbService : DbService
    {
        internal ModuleAutostartDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Пакетоное создание новых сущностей в БД. 
        /// Позиции НЕ дублируются только из-за UNIQUE INDEX в самой БД: (UserId, ProjectId, ModuleId, DBTableName, DBTableKeyId)
        /// </summary>
        public void BulkCreateDBModuleAutostarts(IEnumerable<DBModuleAutostart> dBModuleAutostarts) =>
            ExecuteBulkInsert(
                $"INSERT OR IGNORE INTO {_dbTableName} " +
                    $"({nameof(DBModuleAutostart.UserId)}, {nameof(DBModuleAutostart.RevitVersion)}, {nameof(DBModuleAutostart.ProjectId)}, {nameof(DBModuleAutostart.ModuleId)}, {nameof(DBModuleAutostart.DBTableName)}, {nameof(DBModuleAutostart.DBTableKeyId)}) " +
                    $"VALUES (@{nameof(DBModuleAutostart.UserId)}, @{nameof(DBModuleAutostart.RevitVersion)}, @{nameof(DBModuleAutostart.ProjectId)}, @{nameof(DBModuleAutostart.ModuleId)}, @{nameof(DBModuleAutostart.DBTableName)}, @{nameof(DBModuleAutostart.DBTableKeyId)});",
                dBModuleAutostarts);
        #endregion

        #region Read
        /// <summary>
        /// Получить конфигурации по нужному модулю для нужного пользователя 
        /// </summary>
        public IEnumerable<DBModuleAutostart> GetDBModuleAutostarts_ByUserAndRVersionAndTable(int userId, int rVersion, string tableName) =>
            ExecuteQuery<DBModuleAutostart>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBModuleAutostart.UserId)}='{userId}'" +
                $"AND {nameof(DBModuleAutostart.RevitVersion)}='{rVersion}'" +
                $"AND {nameof(DBModuleAutostart.DBTableName)}='{tableName}';");

        /// <summary>
        /// Получить конфигурации по нужному модулю для нужного пользователя по нужному проекту для нужной таблицы конфигураций
        /// </summary>
        public IEnumerable<DBModuleAutostart> GetDBModuleAutostartsByUserAndRVersionAndPrjIdAndTable(int userId, int rVersion, int prjId, int moduleId, string tableName) =>
            ExecuteQuery<DBModuleAutostart>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBModuleAutostart.UserId)}='{userId}'" +
                $"AND {nameof(DBModuleAutostart.RevitVersion)}='{rVersion}'" +
                $"AND {nameof(DBModuleAutostart.ProjectId)}='{prjId}'" +
                $"AND {nameof(DBModuleAutostart.ModuleId)}='{moduleId}'" +
                $"AND {nameof(DBModuleAutostart.DBTableName)}='{tableName}';");
        #endregion

        #region Update
        #endregion

        #region Delete
        /// <summary>
        /// Удаление сущности в БД
        /// </summary>
        public void DeleteDBModuleAutostarts(DBModuleAutostart dBModuleAutostart) =>
            ExecuteNonQuery(
                $"DELETE FROM {_dbTableName} " +
                $"WHERE {nameof(DBModuleAutostart.UserId)}='{dBModuleAutostart.UserId}'" +
                $"AND {nameof(DBModuleAutostart.ProjectId)}='{dBModuleAutostart.ProjectId}'" +
                $"AND {nameof(DBModuleAutostart.ModuleId)}='{dBModuleAutostart.ModuleId}'" +
                $"AND {nameof(DBModuleAutostart.DBTableName)}='{dBModuleAutostart.DBTableName}'" +
                $"AND {nameof(DBModuleAutostart.DBTableKeyId)}='{dBModuleAutostart.DBTableKeyId}';");
        #endregion
    }
}
