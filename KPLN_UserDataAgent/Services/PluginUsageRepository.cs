using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class PluginUsageRepository
    {
        private const string InvalidLocalSchemaMessagePrefix = "Plugin local database schema is invalid:";
        private const string CentralDatabaseNamePrefix = "KPLN_UserDataAgentPlugin";
        private readonly object _localLock = new object();
        private readonly object _centralSchemaLock = new object();
        private readonly HashSet<string> _ensuredCentralDatabasePaths =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly string _localDatabasePath;
        private readonly string _centralDatabasePath;
        private bool _isLocalSchemaInitialized;
        private int _syncRotationGuard;

        public PluginUsageRepository(string localDatabasePath, string centralDatabasePath)
        {
            _localDatabasePath = localDatabasePath;
            _centralDatabasePath = centralDatabasePath;
        }

        public void InitializeLocal()
        {
            lock (_localLock)
            {
                EnsureDirectory(_localDatabasePath);

                if (Interlocked.CompareExchange(ref _syncRotationGuard, 0, 0) == 0)
                    RotateLocalDatabaseIfNeeded();

                if (_isLocalSchemaInitialized && File.Exists(_localDatabasePath))
                    return;

                InitializeLocalWithRecovery();
            }
        }

        public void InsertEvent(PluginUsageRecord record)
        {
            ExecuteLocalWriteWithRecovery((connection, transaction) =>
                InsertLocalPluginEvent(connection, transaction, record));
        }

        public int TrySyncPendingToCentral(int batchSize)
        {
            InitializeLocal();
            Interlocked.Increment(ref _syncRotationGuard);

            try
            {
                PluginUsageRecord[] pending = LoadPendingBatch(batchSize);
                if (pending.Length == 0)
                {
                    CleanupSyncedArchives();
                    TryPruneCentralDatabases();
                    return 0;
                }

                SyncEventsToCentral(pending);
                MarkEventsSynced(pending);
                CleanupSyncedArchives();
                TryPruneCentralDatabases();
                return pending.Length;
            }
            finally
            {
                Interlocked.Decrement(ref _syncRotationGuard);
            }
        }

        private void InitializeLocalWithRecovery()
        {
            bool databaseFilesExistedBeforeInitialize = LocalDatabaseFilesExist(_localDatabasePath);

            try
            {
                OpenAndEnsureLocalDatabase(databaseFilesExistedBeforeInitialize);
            }
            catch (Exception exception)
            {
                if (!databaseFilesExistedBeforeInitialize || !IsRecoverableLocalDatabaseException(exception))
                    throw;

                RecreateLocalDatabase();
                OpenAndEnsureLocalDatabase(false);
            }
        }

        private void OpenAndEnsureLocalDatabase(bool validateExistingSchema)
        {
            using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
            {
                connection.Open();
                ExecuteNonQuery(connection, null, "PRAGMA journal_mode=WAL;");
                ExecuteNonQuery(connection, null, "PRAGMA busy_timeout=" + ModuleData.LocalBusyTimeoutMs + ";");
                if (validateExistingSchema)
                    ValidateLocalSchema(connection, null);

                EnsureLocalSchema(connection, null);
                ValidateLocalSchema(connection, null);
                _isLocalSchemaInitialized = true;
            }
        }

        private void ExecuteLocalWriteWithRecovery(Action<SQLiteConnection, SQLiteTransaction> writeAction)
        {
            InitializeLocal();

            try
            {
                ExecuteLocalWrite(writeAction);
            }
            catch (Exception exception)
            {
                if (!LocalDatabaseFilesExist(_localDatabasePath) || !IsRecoverableLocalDatabaseException(exception))
                    throw;

                lock (_localLock)
                {
                    RecreateLocalDatabase();
                }

                InitializeLocal();
                ExecuteLocalWrite(writeAction);
            }
        }

        private void ExecuteLocalWrite(Action<SQLiteConnection, SQLiteTransaction> writeAction)
        {
            lock (_localLock)
            {
                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    using (SQLiteTransaction transaction = connection.BeginTransaction())
                    {
                        writeAction(connection, transaction);
                        transaction.Commit();
                    }
                }
            }
        }

        private PluginUsageRecord[] LoadPendingBatch(int batchSize)
        {
            List<PluginUsageRecord> result = new List<PluginUsageRecord>();

            lock (_localLock)
            {
                foreach (string databasePath in GetLocalQueueDatabasePaths())
                {
                    int remaining = batchSize - result.Count;
                    if (remaining <= 0)
                        break;

                    result.AddRange(LoadPendingEvents(databasePath, remaining));
                }
            }

            return result.ToArray();
        }

        private IEnumerable<PluginUsageRecord> LoadPendingEvents(string databasePath, int batchSize)
        {
            List<PluginUsageRecord> result = new List<PluginUsageRecord>();

            try
            {
                using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
                using (SQLiteCommand command = connection.CreateCommand())
                {
                    connection.Open();
                    ValidateLocalSchema(connection, null);
                    EnsureLocalSchema(connection, null);
                    ValidateLocalSchema(connection, null);
                    command.CommandText =
                        "SELECT Id, SyncId, RunId, EventType, EventTime, WindowsUser, DepartmentKey, " +
                        "TabName, PanelName, ButtonName, TransactionName, AddedCount, ModifiedCount, DeletedCount " +
                        "FROM PluginEvents " +
                        "WHERE IsSynced=0 " +
                        "ORDER BY Id " +
                        "LIMIT @Limit;";
                    command.Parameters.AddWithValue("@Limit", batchSize);

                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new PluginUsageRecord
                            {
                                LocalId = Convert.ToInt64(reader["Id"]),
                                LocalDatabasePath = databasePath,
                                SyncId = Convert.ToString(reader["SyncId"]),
                                RunId = Convert.ToString(reader["RunId"]),
                                EventType = Convert.ToString(reader["EventType"]),
                                EventTime = Convert.ToString(reader["EventTime"]),
                                WindowsUser = Convert.ToString(reader["WindowsUser"]),
                                DepartmentKey = Convert.ToString(reader["DepartmentKey"]),
                                TabName = Convert.ToString(reader["TabName"]),
                                PanelName = Convert.ToString(reader["PanelName"]),
                                ButtonName = Convert.ToString(reader["ButtonName"]),
                                TransactionName = Convert.ToString(reader["TransactionName"]),
                                AddedCount = ReadInt(reader, "AddedCount"),
                                ModifiedCount = ReadInt(reader, "ModifiedCount"),
                                DeletedCount = ReadInt(reader, "DeletedCount")
                            });
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                if (TryDeleteRecoverableLocalQueueDatabase(databasePath, exception))
                    return result;

                throw;
            }

            return result;
        }

        private void SyncEventsToCentral(PluginUsageRecord[] records)
        {
            if (records == null || records.Length == 0)
                return;

            foreach (IGrouping<string, PluginUsageRecord> group in records.GroupBy(GetCentralDatabasePath))
            {
                EnsureDirectory(group.Key);
                bool databaseExistedBeforeOpen = File.Exists(group.Key);

                using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.CentralBusyTimeoutMs))
                {
                    connection.Open();
                    ExecuteNonQuery(connection, null, "PRAGMA busy_timeout=" + ModuleData.CentralBusyTimeoutMs + ";");
                    EnsureCentralSchema(connection, null, group.Key, databaseExistedBeforeOpen);

                    using (SQLiteTransaction transaction = connection.BeginTransaction())
                    {
                        foreach (PluginUsageRecord record in group)
                        {
                            InsertCentralPluginEvent(connection, transaction, record);
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        private void MarkEventsSynced(PluginUsageRecord[] records)
        {
            if (records == null || records.Length == 0)
                return;

            lock (_localLock)
            {
                foreach (IGrouping<string, PluginUsageRecord> group in records.GroupBy(e => e.LocalDatabasePath))
                {
                    if (!File.Exists(group.Key))
                        continue;

                    using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.LocalBusyTimeoutMs))
                    {
                        connection.Open();
                        using (SQLiteTransaction transaction = connection.BeginTransaction())
                        {
                            foreach (PluginUsageRecord record in group)
                            {
                                using (SQLiteCommand command = connection.CreateCommand())
                                {
                                    command.Transaction = transaction;
                                    command.CommandText =
                                        "UPDATE PluginEvents " +
                                        "SET IsSynced=1, SyncedAt=@SyncedAt " +
                                        "WHERE Id=@Id;";
                                    command.Parameters.AddWithValue("@SyncedAt", DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"));
                                    command.Parameters.AddWithValue("@Id", record.LocalId);
                                    command.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                        }
                    }
                }
            }
        }

        private string GetCentralDatabasePath(PluginUsageRecord record)
        {
            if (CentralDatabasePathBuilder.IsDatabaseFilePath(_centralDatabasePath))
                return _centralDatabasePath;

            return CentralDatabasePathBuilder.GetDatabasePath(
                _centralDatabasePath,
                record.DepartmentKey,
                record.EventTime,
                CentralDatabaseNamePrefix);
        }

        private void EnsureCentralSchema(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string databasePath,
            bool databaseExistedBeforeOpen)
        {
            if (databaseExistedBeforeOpen && IsCentralSchemaEnsured(databasePath))
                return;

            CreateCentralSchema(connection, transaction);
            MarkCentralSchemaEnsured(databasePath);
        }

        private bool IsCentralSchemaEnsured(string databasePath)
        {
            lock (_centralSchemaLock)
            {
                return _ensuredCentralDatabasePaths.Contains(NormalizeDatabasePath(databasePath));
            }
        }

        private void MarkCentralSchemaEnsured(string databasePath)
        {
            lock (_centralSchemaLock)
            {
                _ensuredCentralDatabasePaths.Add(NormalizeDatabasePath(databasePath));
            }
        }

        private static void EnsureLocalSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            CreatePluginEventsTable(connection, transaction, true);
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_PluginEvents_Sync " +
                "ON PluginEvents(IsSynced, Id);");
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_PluginEvents_Time " +
                "ON PluginEvents(EventTime);");
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_PluginEvents_Run " +
                "ON PluginEvents(RunId);");
        }

        private static void CreateCentralSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            CreatePluginEventsTable(connection, transaction, false);
            EnsureColumn(connection, transaction, "PluginEvents", "PanelName", "TEXT NOT NULL DEFAULT ''");
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_PluginEvents_Time " +
                "ON PluginEvents(EventTime);");
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_PluginEvents_Run " +
                "ON PluginEvents(RunId);");
        }

        private static void CreatePluginEventsTable(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            string localColumns = isLocal
                ? ", IsSynced INTEGER NOT NULL DEFAULT 0, SyncedAt TEXT NOT NULL DEFAULT ''"
                : string.Empty;

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS PluginEvents (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "SyncId TEXT NOT NULL UNIQUE, " +
                "RunId TEXT NOT NULL, " +
                "EventType TEXT NOT NULL, " +
                "EventTime TEXT NOT NULL, " +
                "WindowsUser TEXT NOT NULL, " +
                "DepartmentKey TEXT NOT NULL DEFAULT '" + CentralDatabasePathBuilder.UnknownDepartmentKey + "', " +
                "TabName TEXT NOT NULL, " +
                "PanelName TEXT NOT NULL, " +
                "ButtonName TEXT NOT NULL, " +
                "TransactionName TEXT NOT NULL, " +
                "AddedCount INTEGER NOT NULL DEFAULT 0, " +
                "ModifiedCount INTEGER NOT NULL DEFAULT 0, " +
                "DeletedCount INTEGER NOT NULL DEFAULT 0" +
                localColumns +
                ");");
        }

        private static void ValidateLocalSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            Dictionary<string, LocalTableColumn> actualColumns = ReadTableColumns(connection, transaction, "PluginEvents");
            if (actualColumns.Count == 0)
                ThrowInvalidLocalSchema("Missing table PluginEvents.");

            LocalColumnDefinition[] expectedColumns =
            {
                new LocalColumnDefinition("Id", "INTEGER", false, 1),
                new LocalColumnDefinition("SyncId", "TEXT", true, 0),
                new LocalColumnDefinition("RunId", "TEXT", true, 0),
                new LocalColumnDefinition("EventType", "TEXT", true, 0),
                new LocalColumnDefinition("EventTime", "TEXT", true, 0),
                new LocalColumnDefinition("WindowsUser", "TEXT", true, 0),
                new LocalColumnDefinition("DepartmentKey", "TEXT", true, 0),
                new LocalColumnDefinition("TabName", "TEXT", true, 0),
                new LocalColumnDefinition("PanelName", "TEXT", true, 0),
                new LocalColumnDefinition("ButtonName", "TEXT", true, 0),
                new LocalColumnDefinition("TransactionName", "TEXT", true, 0),
                new LocalColumnDefinition("AddedCount", "INTEGER", true, 0),
                new LocalColumnDefinition("ModifiedCount", "INTEGER", true, 0),
                new LocalColumnDefinition("DeletedCount", "INTEGER", true, 0),
                new LocalColumnDefinition("IsSynced", "INTEGER", true, 0),
                new LocalColumnDefinition("SyncedAt", "TEXT", true, 0)
            };

            Dictionary<string, LocalColumnDefinition> expectedByName =
                expectedColumns.ToDictionary(c => c.Name, StringComparer.OrdinalIgnoreCase);

            foreach (LocalColumnDefinition expectedColumn in expectedColumns)
            {
                LocalTableColumn actualColumn;
                if (!actualColumns.TryGetValue(expectedColumn.Name, out actualColumn))
                    ThrowInvalidLocalSchema("Missing column PluginEvents." + expectedColumn.Name + ".");

                if (!ColumnSchemaMatches(expectedColumn, actualColumn))
                    ThrowInvalidLocalSchema("Column mismatch PluginEvents." + expectedColumn.Name + ".");
            }

            foreach (LocalTableColumn actualColumn in actualColumns.Values)
            {
                if (!expectedByName.ContainsKey(actualColumn.Name))
                    ThrowInvalidLocalSchema("Unexpected column PluginEvents." + actualColumn.Name + ".");
            }
        }

        private static Dictionary<string, LocalTableColumn> ReadTableColumns(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string tableName)
        {
            Dictionary<string, LocalTableColumn> columns =
                new Dictionary<string, LocalTableColumn>(StringComparer.OrdinalIgnoreCase);

            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = "PRAGMA table_info(" + QuoteIdentifier(tableName) + ");";
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        LocalTableColumn column = new LocalTableColumn(
                            Convert.ToString(reader["name"]),
                            Convert.ToString(reader["type"]),
                            ReadPragmaInt(reader, "notnull") != 0,
                            ReadPragmaInt(reader, "pk"));
                        columns[column.Name] = column;
                    }
                }
            }

            return columns;
        }

        private static bool ColumnSchemaMatches(LocalColumnDefinition expected, LocalTableColumn actual)
        {
            return string.Equals(
                    NormalizeColumnType(expected.Type),
                    NormalizeColumnType(actual.Type),
                    StringComparison.OrdinalIgnoreCase)
                && expected.NotNull == actual.NotNull
                && expected.PrimaryKey == actual.PrimaryKey;
        }

        private static string NormalizeColumnType(string columnType)
        {
            return (columnType ?? string.Empty).Trim();
        }

        private static void EnsureColumn(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string tableName,
            string columnName,
            string columnDefinition)
        {
            if (ColumnExists(connection, transaction, tableName, columnName))
                return;

            ExecuteNonQuery(
                connection,
                transaction,
                "ALTER TABLE " + QuoteIdentifier(tableName) +
                " ADD COLUMN " + QuoteIdentifier(columnName) + " " + columnDefinition + ";");
        }

        private static bool ColumnExists(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string tableName,
            string columnName)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = "PRAGMA table_info(" + QuoteIdentifier(tableName) + ");";
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (string.Equals(Convert.ToString(reader["name"]), columnName, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }

            return false;
        }

        private static void InsertLocalPluginEvent(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            PluginUsageRecord record)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO PluginEvents (" +
                    "SyncId, RunId, EventType, EventTime, WindowsUser, DepartmentKey, TabName, PanelName, ButtonName, " +
                    "TransactionName, AddedCount, ModifiedCount, DeletedCount, IsSynced, SyncedAt" +
                    ") VALUES (" +
                    "@SyncId, @RunId, @EventType, @EventTime, @WindowsUser, @DepartmentKey, @TabName, @PanelName, @ButtonName, " +
                    "@TransactionName, @AddedCount, @ModifiedCount, @DeletedCount, 0, ''" +
                    ");";
                AddPluginEventParameters(command, record);
                command.ExecuteNonQuery();
            }
        }

        private static void InsertCentralPluginEvent(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            PluginUsageRecord record)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO PluginEvents (" +
                    "SyncId, RunId, EventType, EventTime, WindowsUser, DepartmentKey, TabName, PanelName, ButtonName, " +
                    "TransactionName, AddedCount, ModifiedCount, DeletedCount" +
                    ") VALUES (" +
                    "@SyncId, @RunId, @EventType, @EventTime, @WindowsUser, @DepartmentKey, @TabName, @PanelName, @ButtonName, " +
                    "@TransactionName, @AddedCount, @ModifiedCount, @DeletedCount" +
                    ");";
                AddPluginEventParameters(command, record);
                command.ExecuteNonQuery();
            }
        }

        private static void AddPluginEventParameters(SQLiteCommand command, PluginUsageRecord record)
        {
            command.Parameters.AddWithValue("@SyncId", record.SyncId);
            command.Parameters.AddWithValue("@RunId", record.RunId ?? string.Empty);
            command.Parameters.AddWithValue("@EventType", record.EventType ?? string.Empty);
            command.Parameters.AddWithValue("@EventTime", record.EventTime ?? string.Empty);
            command.Parameters.AddWithValue("@WindowsUser", record.WindowsUser ?? string.Empty);
            command.Parameters.AddWithValue("@DepartmentKey", CentralDatabasePathBuilder.NormalizeDepartmentKey(record.DepartmentKey));
            command.Parameters.AddWithValue("@TabName", record.TabName ?? string.Empty);
            command.Parameters.AddWithValue("@PanelName", record.PanelName ?? string.Empty);
            command.Parameters.AddWithValue("@ButtonName", record.ButtonName ?? string.Empty);
            command.Parameters.AddWithValue("@TransactionName", record.TransactionName ?? string.Empty);
            command.Parameters.AddWithValue("@AddedCount", record.AddedCount);
            command.Parameters.AddWithValue("@ModifiedCount", record.ModifiedCount);
            command.Parameters.AddWithValue("@DeletedCount", record.DeletedCount);
        }

        private IEnumerable<string> GetLocalQueueDatabasePaths()
        {
            List<string> paths = new List<string>();
            string directoryPath = Path.GetDirectoryName(_localDatabasePath);
            if (!string.IsNullOrWhiteSpace(directoryPath) && Directory.Exists(directoryPath))
            {
                string archivePattern =
                    Path.GetFileNameWithoutExtension(_localDatabasePath) + "_*" + Path.GetExtension(_localDatabasePath);
                paths.AddRange(Directory.GetFiles(directoryPath, archivePattern).OrderBy(p => p, StringComparer.OrdinalIgnoreCase));
            }

            if (File.Exists(_localDatabasePath))
                paths.Add(_localDatabasePath);

            return paths;
        }

        private void RotateLocalDatabaseIfNeeded()
        {
            if (!File.Exists(_localDatabasePath))
                return;

            long maxSizeBytes = (long)ModuleData.LocalDatabaseRotationSizeMb * 1024L * 1024L;
            if (GetDatabaseStorageSize(_localDatabasePath) <= maxSizeBytes)
                return;

            CheckpointDatabase(_localDatabasePath);
            if (GetDatabaseStorageSize(_localDatabasePath) <= maxSizeBytes)
                return;

            DeleteDatabaseSidecars(_localDatabasePath);
            File.Move(_localDatabasePath, CreateArchiveDatabasePath());
            _isLocalSchemaInitialized = false;
        }

        private void CheckpointDatabase(string databasePath)
        {
            using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
            {
                connection.Open();
                ExecuteNonQuery(connection, null, "PRAGMA wal_checkpoint(TRUNCATE);");
            }
        }

        private string CreateArchiveDatabasePath()
        {
            string directoryPath = Path.GetDirectoryName(_localDatabasePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_localDatabasePath);
            string extension = Path.GetExtension(_localDatabasePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string archivePath = Path.Combine(directoryPath, fileNameWithoutExtension + "_" + timestamp + extension);
            int index = 1;

            while (File.Exists(archivePath))
            {
                archivePath = Path.Combine(
                    directoryPath,
                    fileNameWithoutExtension + "_" + timestamp + "_" + index + extension);
                index++;
            }

            return archivePath;
        }

        private void CleanupSyncedArchives()
        {
            lock (_localLock)
            {
                string directoryPath = Path.GetDirectoryName(_localDatabasePath);
                if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
                    return;

                string archivePattern =
                    Path.GetFileNameWithoutExtension(_localDatabasePath) + "_*" + Path.GetExtension(_localDatabasePath);
                foreach (string archivePath in Directory.GetFiles(directoryPath, archivePattern))
                {
                    try
                    {
                        if (!HasPendingRows(archivePath))
                            DeleteDatabaseFiles(archivePath);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private bool HasPendingRows(string databasePath)
        {
            try
            {
                using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    ValidateLocalSchema(connection, null);
                    EnsureLocalSchema(connection, null);
                    ValidateLocalSchema(connection, null);

                    using (SQLiteCommand command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT COUNT(*) FROM PluginEvents WHERE IsSynced=0;";
                        return Convert.ToInt32(command.ExecuteScalar()) > 0;
                    }
                }
            }
            catch (Exception exception)
            {
                if (TryDeleteRecoverableLocalQueueDatabase(databasePath, exception))
                    return false;

                throw;
            }
        }

        private void TryPruneCentralDatabases()
        {
            try
            {
                PruneCentralDatabases();
            }
            catch
            {
            }
        }

        private void PruneCentralDatabases()
        {
            if (string.IsNullOrWhiteSpace(_centralDatabasePath)
                || CentralDatabasePathBuilder.IsDatabaseFilePath(_centralDatabasePath)
                || !Directory.Exists(_centralDatabasePath))
            {
                return;
            }

            int retentionMonths = ModuleData.PluginCentralDatabaseRetentionMonths;
            if (retentionMonths <= 0)
                return;

            DateTime currentMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime cutoffMonth = currentMonth.AddMonths(-(retentionMonths - 1));
            string centralRoot = Path.GetFullPath(_centralDatabasePath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            foreach (string departmentDirectoryPath in Directory.GetDirectories(centralRoot))
            {
                string fullDepartmentDirectoryPath = Path.GetFullPath(departmentDirectoryPath)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (!IsDirectChildPath(centralRoot, fullDepartmentDirectoryPath))
                    continue;

                string departmentKey = Path.GetFileName(fullDepartmentDirectoryPath);
                foreach (string monthDirectoryPath in Directory.GetDirectories(fullDepartmentDirectoryPath))
                {
                    PruneDepartmentMonthDirectory(
                        fullDepartmentDirectoryPath,
                        departmentKey,
                        monthDirectoryPath,
                        cutoffMonth);
                }

                TryDeleteDirectoryIfEmpty(fullDepartmentDirectoryPath);
            }
        }

        private static void PruneDepartmentMonthDirectory(
            string departmentDirectoryPath,
            string departmentKey,
            string monthDirectoryPath,
            DateTime cutoffMonth)
        {
            string monthDirectoryName = Path.GetFileName(monthDirectoryPath);
            DateTime month;
            if (!CentralDatabasePathBuilder.TryParseMonthDirectoryName(monthDirectoryName, out month))
                return;

            if (month >= cutoffMonth)
                return;

            string fullMonthDirectoryPath = Path.GetFullPath(monthDirectoryPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (!IsDirectChildPath(departmentDirectoryPath, fullMonthDirectoryPath))
                return;

            try
            {
                string databasePath = Path.Combine(
                    fullMonthDirectoryPath,
                    CentralDatabasePathBuilder.NormalizeDatabaseNamePrefix(CentralDatabaseNamePrefix)
                    + "_" + CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentKey)
                    + "_" + monthDirectoryName + ".db");
                DeleteDatabaseFiles(databasePath);
                TryDeleteDirectoryIfEmpty(fullMonthDirectoryPath);
            }
            catch
            {
            }
        }

        private void RecreateLocalDatabase()
        {
            _isLocalSchemaInitialized = false;
            DeleteDatabaseFiles(_localDatabasePath);
        }

        private bool TryDeleteRecoverableLocalQueueDatabase(string databasePath, Exception exception)
        {
            if (!IsLocalQueueDatabasePath(databasePath) || !IsRecoverableLocalDatabaseException(exception))
                return false;

            try
            {
                DeleteDatabaseFiles(databasePath);

                if (string.Equals(
                    NormalizeDatabasePath(databasePath),
                    NormalizeDatabasePath(_localDatabasePath),
                    StringComparison.OrdinalIgnoreCase))
                {
                    _isLocalSchemaInitialized = false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsLocalQueueDatabasePath(string databasePath)
        {
            string localPath = NormalizeDatabasePath(_localDatabasePath);
            string targetPath = NormalizeDatabasePath(databasePath);

            string localDirectory = Path.GetDirectoryName(localPath);
            string targetDirectory = Path.GetDirectoryName(targetPath);
            if (string.IsNullOrWhiteSpace(localDirectory)
                || string.IsNullOrWhiteSpace(targetDirectory)
                || !string.Equals(localDirectory, targetDirectory, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string localExtension = Path.GetExtension(localPath);
            string targetExtension = Path.GetExtension(targetPath);
            if (!string.Equals(localExtension, targetExtension, StringComparison.OrdinalIgnoreCase))
                return false;

            string localName = Path.GetFileNameWithoutExtension(localPath);
            string targetName = Path.GetFileNameWithoutExtension(targetPath);
            return string.Equals(localName, targetName, StringComparison.OrdinalIgnoreCase)
                || targetName.StartsWith(localName + "_", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRecoverableLocalDatabaseException(Exception exception)
        {
            if (exception == null)
                return false;

            InvalidOperationException invalidOperationException = exception as InvalidOperationException;
            if (invalidOperationException != null
                && (invalidOperationException.Message ?? string.Empty).StartsWith(
                    InvalidLocalSchemaMessagePrefix,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            SQLiteException sqliteException = exception as SQLiteException;
            if (sqliteException == null)
                return false;

            string message = sqliteException.Message ?? string.Empty;
            return ContainsInvariant(message, "not a database")
                || ContainsInvariant(message, "database disk image is malformed")
                || ContainsInvariant(message, "malformed database schema")
                || ContainsInvariant(message, "schema is corrupt")
                || ContainsInvariant(message, "no such table")
                || ContainsInvariant(message, "no such column")
                || ContainsInvariant(message, "has no column named")
                || ContainsInvariant(message, "is not a table")
                || ContainsInvariant(message, "is a view")
                || ContainsInvariant(message, "cannot modify")
                || ContainsInvariant(message, "NOT NULL constraint failed")
                || ContainsInvariant(message, "datatype mismatch")
                || ContainsInvariant(message, "foreign key mismatch");
        }

        private static bool ContainsInvariant(string source, string value)
        {
            return (source ?? string.Empty).IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool LocalDatabaseFilesExist(string databasePath)
        {
            return File.Exists(databasePath)
                || File.Exists(databasePath + "-wal")
                || File.Exists(databasePath + "-shm");
        }

        private static long GetDatabaseStorageSize(string databasePath)
        {
            long size = GetFileSize(databasePath);
            size += GetFileSize(databasePath + "-wal");
            size += GetFileSize(databasePath + "-shm");
            return size;
        }

        private static long GetFileSize(string path)
        {
            return File.Exists(path) ? new FileInfo(path).Length : 0L;
        }

        private static void DeleteDatabaseFiles(string databasePath)
        {
            DeleteFileIfExists(databasePath);
            DeleteDatabaseSidecars(databasePath);
        }

        private static void DeleteDatabaseSidecars(string databasePath)
        {
            DeleteFileIfExists(databasePath + "-wal");
            DeleteFileIfExists(databasePath + "-shm");
        }

        private static void DeleteFileIfExists(string path)
        {
            if (File.Exists(path))
                File.Delete(path);
        }

        private static bool IsDirectChildPath(string parentPath, string childPath)
        {
            DirectoryInfo parent = new DirectoryInfo(parentPath);
            DirectoryInfo child = new DirectoryInfo(childPath);
            return child.Parent != null
                && string.Equals(
                    parent.FullName.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    child.Parent.FullName.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);
        }

        private static void TryDeleteDirectoryIfEmpty(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath)
                    && !Directory.EnumerateFileSystemEntries(directoryPath).Any())
                {
                    Directory.Delete(directoryPath, false);
                }
            }
            catch
            {
            }
        }

        private static string NormalizeDatabasePath(string databasePath)
        {
            try
            {
                return Path.GetFullPath(databasePath ?? string.Empty);
            }
            catch
            {
                return databasePath ?? string.Empty;
            }
        }

        private static string QuoteIdentifier(string identifier)
        {
            return "\"" + (identifier ?? string.Empty).Replace("\"", "\"\"") + "\"";
        }

        private static int ReadInt(SQLiteDataReader reader, string columnName)
        {
            try
            {
                object value = reader[columnName];
                return value == null || value == DBNull.Value ? 0 : Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static int ReadPragmaInt(SQLiteDataReader reader, string columnName)
        {
            object value = reader[columnName];
            return value == null || value == DBNull.Value ? 0 : Convert.ToInt32(value);
        }

        private static void ThrowInvalidLocalSchema(string message)
        {
            throw new InvalidOperationException(InvalidLocalSchemaMessagePrefix + " " + message);
        }

        private static SQLiteConnection CreateConnection(string databasePath, int busyTimeoutMs)
        {
            return new SQLiteConnection(
                "Data Source=" + databasePath + ";Version=3;BusyTimeout=" + busyTimeoutMs + ";Pooling=False;");
        }

        private static void ExecuteNonQuery(SQLiteConnection connection, SQLiteTransaction transaction, string commandText)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = commandText;
                command.ExecuteNonQuery();
            }
        }

        private static void EnsureDirectory(string databasePath)
        {
            string directoryPath = Path.GetDirectoryName(databasePath);
            if (!string.IsNullOrWhiteSpace(directoryPath))
                Directory.CreateDirectory(directoryPath);
        }

        private sealed class LocalColumnDefinition
        {
            public LocalColumnDefinition(string name, string type, bool notNull, int primaryKey)
            {
                Name = name;
                Type = type;
                NotNull = notNull;
                PrimaryKey = primaryKey;
            }

            public string Name { get; private set; }
            public string Type { get; private set; }
            public bool NotNull { get; private set; }
            public int PrimaryKey { get; private set; }
        }

        private sealed class LocalTableColumn
        {
            public LocalTableColumn(string name, string type, bool notNull, int primaryKey)
            {
                Name = name;
                Type = type;
                NotNull = notNull;
                PrimaryKey = primaryKey;
            }

            public string Name { get; private set; }
            public string Type { get; private set; }
            public bool NotNull { get; private set; }
            public int PrimaryKey { get; private set; }
        }
    }
}