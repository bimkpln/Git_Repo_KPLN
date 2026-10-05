using Dapper;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_DBWorker.Core;
using KPLN_Library_Forms.UI.HtmlWindow;
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

        private readonly string _dbPath;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;

        internal SQLiteService(string dbPath, RevitDocExchangeEnum revitDocExchangeEnum)
        {
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
                case RevitDocExchangeEnum.Revit:
                    ExecuteNonQuery(
                    $"CREATE TABLE {_dbTableName} " +
                        $"({nameof(DBRVTConfigData.Id)} INTEGER PRIMARY KEY, " +
                        $"{nameof(DBRVTConfigData.Name)} TEXT, " +
                        $"{nameof(DBRVTConfigData.PathFrom)} TEXT, " +
                        $"{nameof(DBRVTConfigData.PathTo)} TEXT, " +
                        $"{nameof(DBRVTConfigData.NameChangeFind)} TEXT, " +
                        $"{nameof(DBRVTConfigData.NameChangeSet)} TEXT, " +
                        $"{nameof(DBRVTConfigData.MaxBackup)} INTEGER)");
                    break;
                case RevitDocExchangeEnum.IFC:
                    ExecuteNonQuery(
                    $"CREATE TABLE {_dbTableName} " +
                        $"({nameof(DBIFCConfigData.Id)} INTEGER PRIMARY KEY, " +
                        $"{nameof(DBIFCConfigData.Name)} TEXT, " +
                        $"{nameof(DBIFCConfigData.PathFrom)} TEXT, " +
                        $"{nameof(DBIFCConfigData.PathTo)} TEXT, " +
                        $"{nameof(DBIFCConfigData.FileVersion)} INTEGER, " +
                        $"{nameof(DBIFCConfigData.SpaceBoundaryLevel)} INTEGER, " +
                        $"{nameof(DBIFCConfigData.WallAndColumnSplitting)} INTEGER, " +
                        $"{nameof(DBIFCConfigData.ExportBaseQuantities)} INTEGER, " +
                        $"{nameof(DBIFCConfigData.ExportLinks)} INTEGER, " +
                        $"{nameof(DBIFCConfigData.ViewName)} TEXT, " +
                        $"{nameof(DBIFCConfigData.WorksetToCloseNamesStartWith)} TEXT, " +
                        $"{nameof(DBIFCConfigData.IfcDocPostfix)} TEXT)");
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
            WriteRVTConfigItems(rsConfigs, false);
        }

        internal void ReplaceConfigItems_ByRSConfigs(IEnumerable<DBRVTConfigData> rsConfigs) =>
            WriteRVTConfigItems(rsConfigs, true);

        private void WriteRVTConfigItems(IEnumerable<DBRVTConfigData> rsConfigs, bool replaceExisting)
        {
            DBRVTConfigData[] configs = rsConfigs.ToArray();
            using (IDbConnection connection = new SQLiteConnection(_dbPath))
            {
                connection.Open();
                using (IDbTransaction transaction = connection.BeginTransaction(IsolationLevel.Serializable))
                {
                    string[] columns = GetRVTWritableColumns(connection, transaction, configs);
                    if (replaceExisting)
                        connection.Execute($"DELETE FROM {_dbTableName};", transaction: transaction);

                    connection.Execute(
                        $"INSERT INTO {_dbTableName} ({string.Join(", ", columns)}) " +
                        $"VALUES ({string.Join(", ", columns.Select(column => "@" + column))});",
                        configs, transaction);
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// Добавляем только запрошенную настройку. Старые строки сохраняют MaxBackup = -1.
        /// Отсутствие колонок переименования не мешает сохранить количество резервных копий.
        /// </summary>
        private string[] GetRVTWritableColumns(IDbConnection connection, IDbTransaction transaction, DBRVTConfigData[] configs)
        {
            if (configs.Any(config => config.MaxBackup != -1 && config.MaxBackup <= 0))
                throw new ArgumentOutOfRangeException(nameof(DBRVTConfigData.MaxBackup), "Количество резервных копий должно быть больше нуля.");

            HashSet<string> existingColumns = new HashSet<string>(connection.Query<string>(
                "SELECT name FROM pragma_table_info(@TableName);", new { TableName = _dbTableName }, transaction),
                StringComparer.OrdinalIgnoreCase);

            if (!existingColumns.Contains(nameof(DBRVTConfigData.MaxBackup)) && configs.Any(config => config.MaxBackup != -1))
            {
                connection.Execute($"ALTER TABLE {_dbTableName} ADD COLUMN {nameof(DBRVTConfigData.MaxBackup)} INTEGER DEFAULT -1;", transaction: transaction);
                existingColumns.Add(nameof(DBRVTConfigData.MaxBackup));
            }

            return new[]
            {
                nameof(DBRVTConfigData.Name), nameof(DBRVTConfigData.PathFrom), nameof(DBRVTConfigData.PathTo),
                nameof(DBRVTConfigData.NameChangeFind), nameof(DBRVTConfigData.NameChangeSet), nameof(DBRVTConfigData.MaxBackup),
            }.Where(existingColumns.Contains).ToArray();
        }

        /// <summary>
        /// Создать конфиги IFC
        /// </summary>
        public void PostConfigItems_ByIFCConfigs(IEnumerable<DBIFCConfigData> ifcConfigs)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBIFCConfigData.Name)}, " +
                    $"{nameof(DBIFCConfigData.PathFrom)}, " +
                    $"{nameof(DBIFCConfigData.PathTo)}, " +
                    $"{nameof(DBIFCConfigData.FileVersion)}, " +
                    $"{nameof(DBIFCConfigData.SpaceBoundaryLevel)}, " +
                    $"{nameof(DBIFCConfigData.WallAndColumnSplitting)}, " +
                    $"{nameof(DBIFCConfigData.ExportBaseQuantities)}, " +
                    $"{nameof(DBIFCConfigData.ExportLinks)}, " +
                    $"{nameof(DBIFCConfigData.ViewName)}, " +
                    $"{nameof(DBIFCConfigData.WorksetToCloseNamesStartWith)}, " +
                    $"{nameof(DBIFCConfigData.IfcDocPostfix)}) " +
                $"VALUES " +
                    $"(@{nameof(DBIFCConfigData.Name)}, " +
                    $"@{nameof(DBIFCConfigData.PathFrom)}, " +
                    $"@{nameof(DBIFCConfigData.PathTo)}, " +
                    $"@{nameof(DBIFCConfigData.FileVersion)}, " +
                    $"@{nameof(DBIFCConfigData.SpaceBoundaryLevel)}, " +
                    $"@{nameof(DBIFCConfigData.WallAndColumnSplitting)}, " +
                    $"@{nameof(DBIFCConfigData.ExportBaseQuantities)}, " +
                    $"@{nameof(DBIFCConfigData.ExportLinks)}, " +
                    $"@{nameof(DBIFCConfigData.ViewName)}, " +
                    $"@{nameof(DBIFCConfigData.WorksetToCloseNamesStartWith)}, " +
                    $"@{nameof(DBIFCConfigData.IfcDocPostfix)});",
                ifcConfigs);
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
                case RevitDocExchangeEnum.Revit:
                    return ExecuteQuery<DBRVTConfigData>($"SELECT * FROM {_dbTableName};");
                case RevitDocExchangeEnum.IFC:
                    return ExecuteQuery<DBIFCConfigData>($"SELECT * FROM {_dbTableName};");
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
                case RevitDocExchangeEnum.Revit:
                    if (dBConfig is DBRVTConfigData rsConfig)
                    {
                        using (IDbConnection connection = new SQLiteConnection(_dbPath))
                        {
                            connection.Open();
                            using (IDbTransaction transaction = connection.BeginTransaction(IsolationLevel.Serializable))
                            {
                                string[] columns = GetRVTWritableColumns(connection, transaction, new[] { rsConfig });
                                connection.Execute($"UPDATE {_dbTableName} SET " +
                                    string.Join(", ", columns.Select(column => column + " = @" + column)) + " WHERE Id = @Id;",
                                    rsConfig, transaction);
                                DBRVTConfigData result = connection.Query<DBRVTConfigData>(
                                    $"SELECT * FROM {_dbTableName} WHERE Id = @Id;", rsConfig, transaction).FirstOrDefault();
                                transaction.Commit();
                                return result;
                            }
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
