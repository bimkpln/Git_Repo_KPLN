using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using KPLN_Loader.Core.Entities;
using System.Linq;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом PluginActivity в БД
    /// </summary>
    public class SQLitePluginActivityService : SQLiteService
    {
        internal SQLitePluginActivityService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Создать новый PluginActivity в БД
        /// </summary>
        public void CreateDBPluginActivity(DBPluginActivity dBPluginActivity)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBPluginActivity.ModuleId)}, " +
                    $"{nameof(DBPluginActivity.PluginName)}, " +
                    $"{nameof(DBPluginActivity.UserId)}, " +
                    $"{nameof(DBPluginActivity.UsageCount)}, " +
                    $"{nameof(DBPluginActivity.LastActivityDate)}) " +
                $"VALUES (" +
                    $"@{nameof(DBPluginActivity.ModuleId)}, " +
                    $"@{nameof(DBPluginActivity.PluginName)}, " +
                    $"@{nameof(DBPluginActivity.UserId)}, " +
                    $"@{nameof(DBPluginActivity.UsageCount)}, " +
                    $"@{nameof(DBPluginActivity.LastActivityDate)});",
            dBPluginActivity);
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить PluginActivity по имени плагина и id пользователя
        /// </summary>
        public DBPluginActivity GetDBPluginActivity_ByModuleNameAndSubDep(string pluginName, int userId) =>
            ExecuteQuery<DBPluginActivity>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBPluginActivity.PluginName)}='{pluginName}' AND {nameof(DBPluginActivity.UserId)}='{userId}';")
            .FirstOrDefault();
        #endregion

        #region Update
        /// <summary>
        /// Обновить значение PluginActivity для модуля по DBPluginActivity и id пользователя
        /// </summary>
        public void UpdatePluginActivity_ByPluginActivityAndSubDep(DBPluginActivity dBPluginActivity, int userId)
        {
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                    $"SET {nameof(DBPluginActivity.UsageCount)}='{dBPluginActivity.UsageCount}', " +
                    $"{nameof(DBPluginActivity.LastActivityDate)}='{dBPluginActivity.LastActivityDate}' " +
                    $"WHERE {nameof(DBPluginActivity.Id)}='{dBPluginActivity.Id}' AND {nameof(DBPluginActivity.UserId)}='{userId}';");
            #endregion
        }
    }
}
