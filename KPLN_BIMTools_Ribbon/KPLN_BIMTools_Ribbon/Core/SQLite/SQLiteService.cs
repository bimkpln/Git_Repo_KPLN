using Dapper;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.Core.SQLite
{
    /// <summary>
    /// Сервис для работы с БД
    /// </summary>
    public class SQLiteService
    {
        private const string _dbTableName = "Items";

        private readonly Logger _logger;
        private readonly string _dbPath;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;

        internal SQLiteService(Logger logger, string dbPath, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            _logger = logger;
            CurrentDBFullPath = dbPath;
            _dbPath = "Data Source=" + CurrentDBFullPath + "; Version=3;";
            _revitDocExchangeEnum = revitDocExchangeEnum;
        }

        public string CurrentDBFullPath { get; }

        #region Create
        // <summary>
        /// Создать конфиг
        /// </summary>
        internal void CreateDbFile()
        {
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    ExecuteNonQuery(
                    $"CREATE TABLE {_dbTableName} " +
                        $"({nameof(DBNWConfigData.Id)} INTEGER PRIMARY KEY, " +
                        $"{nameof(DBNWConfigData.Name)} TEXT, " +
                        $"{nameof(DBNWConfigData.PathFrom)} TEXT, " +
                        $"{nameof(DBNWConfigData.PathTo)} TEXT, " +
                        $"{nameof(DBNWConfigData.FacetingFactor)} REAL, " +
                        $"{nameof(DBNWConfigData.ConvertElementProperties)} INTEGER, " +
                        $"{nameof(DBNWConfigData.ExportLinks)} INTEGER, " +
                        $"{nameof(DBNWConfigData.FindMissingMaterials)} INTEGER, " +
                        $"{nameof(DBNWConfigData.ExportScope)} INTEGER, " +
                        $"{nameof(DBNWConfigData.DivideFileIntoLevels)} INTEGER, " +
                        $"{nameof(DBNWConfigData.ExportRoomGeometry)} INTEGER, " +
                        $"{nameof(DBNWConfigData.ViewName)} TEXT, " +
                        $"{nameof(DBNWConfigData.WorksetToCloseNamesStartWith)} TEXT, " +
                        $"{nameof(DBNWConfigData.NavisDocPostfix)} TEXT)");
                    break;
                case RevitDocExchangeEnum.RevitServer:
                    ExecuteNonQuery(
                    $"CREATE TABLE {_dbTableName} " +
                        $"({nameof(DBRVTConfigData.Id)} INTEGER PRIMARY KEY, " +
                        $"{nameof(DBRVTConfigData.Name)} TEXT, " +
                        $"{nameof(DBRVTConfigData.PathFrom)} TEXT, " +
                        $"{nameof(DBRVTConfigData.PathTo)} TEXT, " +
                        $"{nameof(DBRVTConfigData.MaxBackup)} INTEGER)");
                    break;
            }
        }

        /// <summary>
        /// Создать конфига
        /// </summary>
        public void PostConfigItems_ByNWConfigs(IEnumerable<DBNWConfigData> nwConfigs)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBNWConfigData.Name)}, " +
                    $"{nameof(DBNWConfigData.PathFrom)}, " +
                    $"{nameof(DBNWConfigData.PathTo)}, " +
                    $"{nameof(DBNWConfigData.FacetingFactor)}, " +
                    $"{nameof(DBNWConfigData.ConvertElementProperties)}, " +
                    $"{nameof(DBNWConfigData.ExportLinks)}, " +
                    $"{nameof(DBNWConfigData.FindMissingMaterials)}, " +
                    $"{nameof(DBNWConfigData.ExportScope)}, " +
                    $"{nameof(DBNWConfigData.DivideFileIntoLevels)}, " +
                    $"{nameof(DBNWConfigData.ExportRoomGeometry)}, " +
                    $"{nameof(DBNWConfigData.ViewName)}, " +
                    $"{nameof(DBNWConfigData.WorksetToCloseNamesStartWith)}, " +
                    $"{nameof(DBNWConfigData.NavisDocPostfix)}) " +
                $"VALUES " +
                    $"(@{nameof(DBNWConfigData.Name)}, " +
                    $"@{nameof(DBNWConfigData.PathFrom)}, " +
                    $"@{nameof(DBNWConfigData.PathTo)}, " +
                    $"@{nameof(DBNWConfigData.FacetingFactor)}, " +
                    $"@{nameof(DBNWConfigData.ConvertElementProperties)}, " +
                    $"@{nameof(DBNWConfigData.ExportLinks)}, " +
                    $"@{nameof(DBNWConfigData.FindMissingMaterials)}, " +
                    $"@{nameof(DBNWConfigData.ExportScope)}, " +
                    $"@{nameof(DBNWConfigData.DivideFileIntoLevels)}, " +
                    $"@{nameof(DBNWConfigData.ExportRoomGeometry)}, " +
                    $"@{nameof(DBNWConfigData.ViewName)}, " +
                    $"@{nameof(DBNWConfigData.WorksetToCloseNamesStartWith)}, " +
                    $"@{nameof(DBNWConfigData.NavisDocPostfix)});",
                nwConfigs);
        }

        /// <summary>
        /// Создать конфига
        /// </summary>
        public void PostConfigItems_ByRSConfigs(IEnumerable<DBRVTConfigData> rsConfigs)
        {
            try
            {
                ExecuteNonQuery(
                    $"INSERT INTO {_dbTableName} " +
                        $"({nameof(DBRVTConfigData.Name)}, " +
                        $"{nameof(DBRVTConfigData.PathFrom)}, " +
                        $"{nameof(DBRVTConfigData.PathTo)}, " +
                        $"{nameof(DBRVTConfigData.MaxBackup)}) " +
                    $"VALUES " +
                        $"(@{nameof(DBRVTConfigData.Name)}, " +
                        $"@{nameof(DBRVTConfigData.PathFrom)}, " +
                        $"@{nameof(DBRVTConfigData.PathTo)}, " +
                        $"@{nameof(DBRVTConfigData.MaxBackup)});",
                    rsConfigs);
            }
            // Старая версия БД, когда не было параметра кол-ва рез. копий
            catch (Exception ex)
            {
                HtmlOutput.Print(
                    "Не удалось перезаписать параметр кол-ва резервных копий, оно останется пустым (дефолтным). " +
                    "Если нужно его заменить - сними копию текушего конфига, а старую удали",
                    MessageType.Warning);

                ExecuteNonQuery(
                   $"INSERT INTO {_dbTableName} " +
                       $"({nameof(DBRVTConfigData.Name)}, " +
                       $"{nameof(DBRVTConfigData.PathFrom)}, " +
                       $"{nameof(DBRVTConfigData.PathTo)}) " +
                   $"VALUES " +
                       $"(@{nameof(DBRVTConfigData.Name)}, " +
                       $"@{nameof(DBRVTConfigData.PathFrom)}, " +
                       $"@{nameof(DBRVTConfigData.PathTo)});",
                   rsConfigs);
            }
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить конфиги по проекту
        /// </summary>
        internal IEnumerable<DBConfigEntity> GetConfigItems()
        {
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    return ExecuteQuery<DBNWConfigData>($"SELECT * FROM {_dbTableName};");
                case RevitDocExchangeEnum.RevitServer:
                    return ExecuteQuery<DBRVTConfigData>($"SELECT * FROM {_dbTableName};");
            }

            return null;
        }
        #endregion

        #region Update
        /// <summary>
        /// Обновить настройки конфига
        /// </summary>
        public DBConfigEntity UpdateConfigItems_ByConfig(DBConfigEntity dBConfig)
        {
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    if (dBConfig is DBNWConfigData nwConfig)
                    {
                        return ExecuteQuery<DBNWConfigData>(
                            $"UPDATE {_dbTableName} " +
                            $"SET " +
                                $"{nameof(DBNWConfigData.PathFrom)} = '{nwConfig.PathFrom}'" +
                                $"{nameof(DBNWConfigData.PathTo)} = '{nwConfig.PathTo}'" +
                                $"{nameof(DBNWConfigData.FacetingFactor)} = '{nwConfig.FacetingFactor}'" +
                                $"{nameof(DBNWConfigData.ConvertElementProperties)} = '{nwConfig.ConvertElementProperties}'" +
                                $"{nameof(DBNWConfigData.ExportLinks)} = '{nwConfig.ExportLinks}'" +
                                $"{nameof(DBNWConfigData.FindMissingMaterials)} = '{nwConfig.FindMissingMaterials}'" +
                                $"{nameof(DBNWConfigData.ExportScope)} = '{nwConfig.ExportScope}'" +
                                $"{nameof(DBNWConfigData.ExportRoomGeometry)} = '{nwConfig.ExportRoomGeometry}'" +
                                $"{nameof(DBNWConfigData.ViewName)} = '{nwConfig.ViewName}'" +
                                $"{nameof(DBNWConfigData.WorksetToCloseNamesStartWith)} = '{nwConfig.WorksetToCloseNamesStartWith}'" +
                                $"{nameof(DBNWConfigData.NavisDocPostfix)} = '{nwConfig.NavisDocPostfix}'" +
                            $"WHERE " +
                                $"{nameof(DBNWConfigData.Id)} = '{nwConfig.Id}';",
                            nwConfig)
                            .FirstOrDefault();
                    }
                    return null;
                case RevitDocExchangeEnum.RevitServer:
                    if (dBConfig is DBRVTConfigData rsConfig)
                    {
                        try
                        {
                            return ExecuteQuery<DBRVTConfigData>(
                                $"UPDATE {_dbTableName} " +
                                $"SET " +
                                    $"{nameof(DBRVTConfigData.PathFrom)} = '{rsConfig.PathFrom}'" +
                                    $"{nameof(DBRVTConfigData.PathTo)} = '{rsConfig.PathTo}'" +
                                    $"{nameof(DBRVTConfigData.MaxBackup)} = '{rsConfig.MaxBackup}'" +
                                $"WHERE " +
                                    $"{nameof(DBRVTConfigData.Id)} = '{rsConfig.Id}';",
                                rsConfig)
                                .FirstOrDefault();
                        }
                        // Старая версия БД, когда не было параметра кол-ва рез. копий
                        catch (Exception ex)
                        {
                            HtmlOutput.Print(
                                "Не удалось перезаписать параметр кол-ва резервных копий, оно останется пустым (дефолтным). " +
                                "Если нужно его заменить - сними копию текушего конфига, а старую удали", 
                                MessageType.Warning);
                            
                            return ExecuteQuery<DBRVTConfigData>(
                                $"UPDATE {_dbTableName} " +
                                $"SET " +
                                    $"{nameof(DBRVTConfigData.PathFrom)} = '{rsConfig.PathFrom}'" +
                                    $"{nameof(DBRVTConfigData.PathTo)} = '{rsConfig.PathTo}'" +
                                $"WHERE " +
                                    $"{nameof(DBRVTConfigData.Id)} = '{rsConfig.Id}';",
                                rsConfig)
                                .FirstOrDefault();
                        }
                    }
                    return null;
            }

            return null;
        }
        #endregion

        #region Delete
        /// <summary>
        /// Очистить таблицу от всех данных
        /// </summary>
        public void DropTable() 
        {
            ExecuteNonQuery($"DELETE FROM {_dbTableName};");
        }
        #endregion

        private void ExecuteNonQuery(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_dbPath))
            {
                connection.Open();
                connection.Execute(query, parameters);
            }
        }

        private IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            using (IDbConnection connection = new SQLiteConnection(_dbPath))
            {
                connection.Open();
                return connection.Query<T>(query, parameters);
            }
        }
    }
}
