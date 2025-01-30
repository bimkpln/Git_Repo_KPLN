using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом PluginActivity в БД
    /// </summary>
    public class PluginActivityDbService : DbService
    {
        internal PluginActivityDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
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
                    $"{nameof(DBPluginActivity.SubDepartmentId)}, " +
                    $"{nameof(DBPluginActivity.UsageCount)}, " +
                    $"{nameof(DBPluginActivity.LastActivityDate)}) " +
                $"VALUES (" +
                    $"@{nameof(DBPluginActivity.ModuleId)}, " +
                    $"@{nameof(DBPluginActivity.PluginName)}, " +
                    $"@{nameof(DBPluginActivity.SubDepartmentId)}, " +
                    $"@{nameof(DBPluginActivity.UsageCount)}, " +
                    $"@{nameof(DBPluginActivity.LastActivityDate)});",
            dBPluginActivity);
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить PluginActivity по имени и отделу
        /// </summary>
        public DBPluginActivity GetDBPluginActivity_ByModuleNameAndSubDep(string pluginName, int subDepId) =>
            ExecuteQuery<DBPluginActivity>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBPluginActivity.PluginName)}='{pluginName}' AND {nameof(DBPluginActivity.SubDepartmentId)}='{subDepId}';")
            .FirstOrDefault();
        #endregion

        #region Update
        /// <summary>
        /// Обновить значение PluginActivity для модуля по DBPluginActivity и отделу
        /// </summary>
        public void UpdatePluginActivity_ByPluginActivityAndSubDep(DBPluginActivity dBPluginActivity, int subDepId)
        {
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                    $"SET {nameof(DBPluginActivity.UsageCount)}='{dBPluginActivity.UsageCount}', " +
                    $"{nameof(DBPluginActivity.LastActivityDate)}='{dBPluginActivity.LastActivityDate}' " +
                    $"WHERE {nameof(DBPluginActivity.Id)}='{dBPluginActivity.Id}' AND {nameof(DBPluginActivity.SubDepartmentId)}='{subDepId}';");
            #endregion
        }
    }
}
