using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Threading;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class UserDataRepository
    {
        private readonly object _localLock = new object();
        private readonly object _centralSchemaLock = new object();
        private readonly HashSet<string> _ensuredCentralDatabasePaths =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly string _localDatabasePath;
        private readonly string _centralDatabasePath;
        private bool _isLocalSchemaInitialized;
        private int _syncRotationGuard;

        public UserDataRepository(string localDatabasePath, string centralDatabasePath)
        {
            _localDatabasePath = localDatabasePath;
            _centralDatabasePath = centralDatabasePath;
        }

        public UserDataRepository(string centralDatabasePath)
            : this(ModuleData.LocalDatabasePath, centralDatabasePath)
        {
        }

        public void Initialize()
        {
            InitializeLocal();
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

                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    ExecuteNonQuery(connection, null, "PRAGMA journal_mode=WAL;");
                    ExecuteNonQuery(connection, null, "PRAGMA busy_timeout=" + ModuleData.LocalBusyTimeoutMs + ";");
                    EnsureLocalSchema(connection, null);
                    _isLocalSchemaInitialized = true;
                }
            }
        }

        public void InsertEvent(UserEventRecord userEvent)
        {
            InitializeLocal();

            lock (_localLock)
            {
                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    using (SQLiteTransaction transaction = connection.BeginTransaction())
                    {
                        InsertLocalUserEvent(connection, transaction, userEvent);
                        transaction.Commit();
                    }
                }
            }
        }

        public void InsertError(string source, Exception exception)
        {
            if (exception == null)
                return;

            AgentErrorRecord error = AgentErrorRecord.Create(source, exception);
            InitializeLocal();

            lock (_localLock)
            {
                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    using (SQLiteTransaction transaction = connection.BeginTransaction())
                    {
                        InsertLocalAgentError(connection, transaction, error);
                        transaction.Commit();
                    }
                }
            }
        }

        public int TrySyncPendingToCentral(int batchSize)
        {
            InitializeLocal();
            Interlocked.Increment(ref _syncRotationGuard);

            try
            {
                PendingSyncBatch pending = LoadPendingBatch(batchSize);
                if (!pending.HasRows)
                {
                    CleanupSyncedArchives();
                    TryPruneCentralDatabases();
                    return 0;
                }

                RefreshUnknownDepartmentKeys(pending);

                int syncedCount = 0;
                Exception syncException = null;

                try
                {
                    SyncUserEventsToCentral(pending.UserEvents);
                    MarkEventsSynced(pending.UserEvents);
                    syncedCount += pending.UserEvents.Length;
                }
                catch (Exception exception)
                {
                    syncException = exception;
                }

                try
                {
                    SyncErrorsToCentral(pending.Errors);
                    MarkErrorsSynced(pending.Errors);
                    syncedCount += pending.Errors.Length;
                }
                catch (Exception exception)
                {
                    if (syncException == null)
                        syncException = exception;
                }

                CleanupSyncedArchives();
                TryPruneCentralDatabases();

                if (syncException != null)
                    throw syncException;

                return syncedCount;
            }
            finally
            {
                Interlocked.Decrement(ref _syncRotationGuard);
            }
        }

        private void SyncUserEventsToCentral(UserEventRecord[] userEvents)
        {
            if (userEvents == null || userEvents.Length == 0)
                return;

            foreach (IGrouping<string, UserEventRecord> group in userEvents.GroupBy(GetCentralEventDatabasePath))
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
                        foreach (UserEventRecord userEvent in group)
                        {
                            InsertCentralUserEvent(connection, transaction, userEvent);
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        private void SyncErrorsToCentral(AgentErrorRecord[] errors)
        {
            if (errors == null || errors.Length == 0)
                return;

            foreach (IGrouping<string, AgentErrorRecord> group in errors.GroupBy(GetCentralErrorDatabasePath))
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
                        foreach (AgentErrorRecord error in group)
                        {
                            InsertCentralAgentError(connection, transaction, error);
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        private PendingSyncBatch LoadPendingBatch(int batchSize)
        {
            List<UserEventRecord> userEvents = new List<UserEventRecord>();
            List<AgentErrorRecord> errors = new List<AgentErrorRecord>();

            lock (_localLock)
            {
                foreach (string databasePath in GetLocalQueueDatabasePaths())
                {
                    int remainingEvents = batchSize - userEvents.Count;
                    if (remainingEvents > 0)
                        userEvents.AddRange(LoadPendingEvents(databasePath, remainingEvents));

                    int remainingErrors = batchSize - errors.Count;
                    if (remainingErrors > 0)
                        errors.AddRange(LoadPendingErrors(databasePath, remainingErrors));

                    if (userEvents.Count >= batchSize && errors.Count >= batchSize)
                        break;
                }
            }

            return new PendingSyncBatch(userEvents.ToArray(), errors.ToArray());
        }

        private IEnumerable<UserEventRecord> LoadPendingEvents(string databasePath, int batchSize)
        {
            List<UserEventRecord> result = new List<UserEventRecord>();
            using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
            using (SQLiteCommand command = connection.CreateCommand())
            {
                connection.Open();
                EnsureLocalSchema(connection, null);
                command.CommandText =
                    "SELECT ue.Id, ue.SyncId, ue.EventTime, ue.WindowsUser, ue.DepartmentKey, ue.EventTransactionId, " +
                    "et.EventName, et.TransactionName, " +
                    "ue.AddedCount, ue.ModifiedCount, ue.DeletedCount " +
                    "FROM UserEvents ue " +
                    "LEFT JOIN EventTransactions et ON et.Id=ue.EventTransactionId " +
                    "WHERE ue.IsSynced=0 " +
                    "ORDER BY ue.Id " +
                    "LIMIT @Limit;";
                command.Parameters.AddWithValue("@Limit", batchSize);

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new UserEventRecord
                        {
                            LocalId = Convert.ToInt64(reader["Id"]),
                            LocalDatabasePath = databasePath,
                            SyncId = Convert.ToString(reader["SyncId"]),
                            EventTime = Convert.ToString(reader["EventTime"]),
                            WindowsUser = Convert.ToString(reader["WindowsUser"]),
                            DepartmentKey = Convert.ToString(reader["DepartmentKey"]),
                            EventTransactionId = Convert.ToInt64(reader["EventTransactionId"]),
                            EventName = Convert.ToString(reader["EventName"]),
                            TransactionName = Convert.ToString(reader["TransactionName"]),
                            AddedCount = ReadInt(reader, "AddedCount"),
                            ModifiedCount = ReadInt(reader, "ModifiedCount"),
                            DeletedCount = ReadInt(reader, "DeletedCount")
                        });
                    }
                }
            }

            return result;
        }

        private IEnumerable<AgentErrorRecord> LoadPendingErrors(string databasePath, int batchSize)
        {
            List<AgentErrorRecord> result = new List<AgentErrorRecord>();
            using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
            using (SQLiteCommand command = connection.CreateCommand())
            {
                connection.Open();
                EnsureLocalSchema(connection, null);
                command.CommandText =
                    "SELECT Id, SyncId, ErrorTime, WindowsUser, DepartmentKey, Source, ErrorType, ErrorMessage, ErrorStackTrace " +
                    "FROM AgentErrors " +
                    "WHERE IsSynced=0 " +
                    "ORDER BY Id " +
                    "LIMIT @Limit;";
                command.Parameters.AddWithValue("@Limit", batchSize);

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new AgentErrorRecord
                        {
                            LocalId = Convert.ToInt64(reader["Id"]),
                            LocalDatabasePath = databasePath,
                            SyncId = Convert.ToString(reader["SyncId"]),
                            ErrorTime = Convert.ToString(reader["ErrorTime"]),
                            WindowsUser = Convert.ToString(reader["WindowsUser"]),
                            DepartmentKey = Convert.ToString(reader["DepartmentKey"]),
                            Source = Convert.ToString(reader["Source"]),
                            ErrorType = Convert.ToString(reader["ErrorType"]),
                            ErrorMessage = Convert.ToString(reader["ErrorMessage"]),
                            ErrorStackTrace = Convert.ToString(reader["ErrorStackTrace"])
                        });
                    }
                }
            }

            return result;
        }

        private void RefreshUnknownDepartmentKeys(PendingSyncBatch pending)
        {
            List<UserEventRecord> updatedEvents = new List<UserEventRecord>();
            List<AgentErrorRecord> updatedErrors = new List<AgentErrorRecord>();
            bool forceReferenceRefresh = true;

            foreach (UserEventRecord userEvent in pending.UserEvents)
            {
                bool wasUnknownDepartment = IsUnknownDepartmentKey(userEvent.DepartmentKey);
                string resolvedDepartmentKey;
                if (TryResolveUnknownDepartmentKey(
                    userEvent.WindowsUser,
                    userEvent.DepartmentKey,
                    wasUnknownDepartment && forceReferenceRefresh,
                    out resolvedDepartmentKey))
                {
                    userEvent.DepartmentKey = resolvedDepartmentKey;
                    updatedEvents.Add(userEvent);
                }

                if (wasUnknownDepartment)
                    forceReferenceRefresh = false;
            }

            foreach (AgentErrorRecord error in pending.Errors)
            {
                bool wasUnknownDepartment = IsUnknownDepartmentKey(error.DepartmentKey);
                string resolvedDepartmentKey;
                if (TryResolveUnknownDepartmentKey(
                    error.WindowsUser,
                    error.DepartmentKey,
                    wasUnknownDepartment && forceReferenceRefresh,
                    out resolvedDepartmentKey))
                {
                    error.DepartmentKey = resolvedDepartmentKey;
                    updatedErrors.Add(error);
                }

                if (wasUnknownDepartment)
                    forceReferenceRefresh = false;
            }

            TryUpdateLocalEventDepartmentKeys(updatedEvents.ToArray());
            TryUpdateLocalErrorDepartmentKeys(updatedErrors.ToArray());
        }

        private static bool TryResolveUnknownDepartmentKey(
            string windowsUser,
            string currentDepartmentKey,
            bool forceReferenceRefresh,
            out string resolvedDepartmentKey)
        {
            resolvedDepartmentKey = currentDepartmentKey;
            if (!IsUnknownDepartmentKey(currentDepartmentKey))
                return false;

            string departmentKey = ReferenceDepartmentLookupService.ResolveDepartmentKey(
                windowsUser,
                ModuleData.ReferenceDatabasePath,
                forceReferenceRefresh);
            if (IsUnknownDepartmentKey(departmentKey))
                return false;

            resolvedDepartmentKey = CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentKey);
            return true;
        }

        private static bool IsUnknownDepartmentKey(string departmentKey)
        {
            return string.Equals(
                CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentKey),
                CentralDatabasePathBuilder.UnknownDepartmentKey,
                StringComparison.OrdinalIgnoreCase);
        }

        private void TryUpdateLocalEventDepartmentKeys(UserEventRecord[] userEvents)
        {
            if (userEvents == null || userEvents.Length == 0)
                return;

            try
            {
                lock (_localLock)
                {
                    foreach (IGrouping<string, UserEventRecord> group in userEvents.GroupBy(e => e.LocalDatabasePath))
                    {
                        using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.LocalBusyTimeoutMs))
                        {
                            connection.Open();
                            using (SQLiteTransaction transaction = connection.BeginTransaction())
                            {
                                foreach (UserEventRecord userEvent in group)
                                {
                                    using (SQLiteCommand command = connection.CreateCommand())
                                    {
                                        command.Transaction = transaction;
                                        command.CommandText =
                                            "UPDATE UserEvents " +
                                            "SET DepartmentKey=@DepartmentKey " +
                                            "WHERE Id=@Id AND DepartmentKey=@UnknownDepartmentKey;";
                                        command.Parameters.AddWithValue(
                                            "@DepartmentKey",
                                            CentralDatabasePathBuilder.NormalizeDepartmentKey(userEvent.DepartmentKey));
                                        command.Parameters.AddWithValue("@Id", userEvent.LocalId);
                                        command.Parameters.AddWithValue(
                                            "@UnknownDepartmentKey",
                                            CentralDatabasePathBuilder.UnknownDepartmentKey);
                                        command.ExecuteNonQuery();
                                    }
                                }

                                transaction.Commit();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void TryUpdateLocalErrorDepartmentKeys(AgentErrorRecord[] errors)
        {
            if (errors == null || errors.Length == 0)
                return;

            try
            {
                lock (_localLock)
                {
                    foreach (IGrouping<string, AgentErrorRecord> group in errors.GroupBy(e => e.LocalDatabasePath))
                    {
                        using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.LocalBusyTimeoutMs))
                        {
                            connection.Open();
                            using (SQLiteTransaction transaction = connection.BeginTransaction())
                            {
                                foreach (AgentErrorRecord error in group)
                                {
                                    using (SQLiteCommand command = connection.CreateCommand())
                                    {
                                        command.Transaction = transaction;
                                        command.CommandText =
                                            "UPDATE AgentErrors " +
                                            "SET DepartmentKey=@DepartmentKey " +
                                            "WHERE Id=@Id AND DepartmentKey=@UnknownDepartmentKey;";
                                        command.Parameters.AddWithValue(
                                            "@DepartmentKey",
                                            CentralDatabasePathBuilder.NormalizeDepartmentKey(error.DepartmentKey));
                                        command.Parameters.AddWithValue("@Id", error.LocalId);
                                        command.Parameters.AddWithValue(
                                            "@UnknownDepartmentKey",
                                            CentralDatabasePathBuilder.UnknownDepartmentKey);
                                        command.ExecuteNonQuery();
                                    }
                                }

                                transaction.Commit();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void MarkEventsSynced(UserEventRecord[] userEvents)
        {
            if (userEvents == null || userEvents.Length == 0)
                return;

            lock (_localLock)
            {
                foreach (IGrouping<string, UserEventRecord> group in userEvents.GroupBy(e => e.LocalDatabasePath))
                {
                    using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.LocalBusyTimeoutMs))
                    {
                        connection.Open();
                        using (SQLiteTransaction transaction = connection.BeginTransaction())
                        {
                            foreach (UserEventRecord userEvent in group)
                            {
                                using (SQLiteCommand command = connection.CreateCommand())
                                {
                                    command.Transaction = transaction;
                                    command.CommandText =
                                        "UPDATE UserEvents " +
                                        "SET IsSynced=1, SyncedAt=@SyncedAt " +
                                        "WHERE Id=@Id;";
                                    command.Parameters.AddWithValue("@SyncedAt", DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"));
                                    command.Parameters.AddWithValue("@Id", userEvent.LocalId);
                                    command.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                        }
                    }
                }
            }
        }

        private void MarkErrorsSynced(AgentErrorRecord[] errors)
        {
            if (errors == null || errors.Length == 0)
                return;

            lock (_localLock)
            {
                foreach (IGrouping<string, AgentErrorRecord> group in errors.GroupBy(e => e.LocalDatabasePath))
                {
                    using (SQLiteConnection connection = CreateConnection(group.Key, ModuleData.LocalBusyTimeoutMs))
                    {
                        connection.Open();
                        using (SQLiteTransaction transaction = connection.BeginTransaction())
                        {
                            foreach (AgentErrorRecord error in group)
                            {
                                using (SQLiteCommand command = connection.CreateCommand())
                                {
                                    command.Transaction = transaction;
                                    command.CommandText =
                                        "UPDATE AgentErrors " +
                                        "SET IsSynced=1, SyncedAt=@SyncedAt " +
                                        "WHERE Id=@Id;";
                                    command.Parameters.AddWithValue("@SyncedAt", DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"));
                                    command.Parameters.AddWithValue("@Id", error.LocalId);
                                    command.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                        }
                    }
                }
            }
        }

        private static void EnsureLocalSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            EnsureEventTransactionsSchema(connection, transaction);
            EnsureUserEventsSchema(connection, transaction, true);
            EnsureAgentErrorsSchema(connection, transaction, true);
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

        private static void CreateCentralSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            EnsureEventTransactionsSchema(connection, transaction);
            EnsureUserEventsSchema(connection, transaction, false);
            EnsureAgentErrorsSchema(connection, transaction, false);
        }

        private static void EnsureUserEventsSchema(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            CreateUserEventsTable(connection, transaction, isLocal);
            if (isLocal)
            {
                EnsureColumn(
                    connection,
                    transaction,
                    "UserEvents",
                    "DepartmentKey",
                    "TEXT NOT NULL DEFAULT '" + CentralDatabasePathBuilder.UnknownDepartmentKey + "'");
            }

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_UserEvents_Time " +
                "ON UserEvents(EventTime);");

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_UserEvents_EventTransaction " +
                "ON UserEvents(EventTransactionId);");

            if (isLocal)
            {
                ExecuteNonQuery(
                    connection,
                    transaction,
                    "CREATE INDEX IF NOT EXISTS IX_UserEvents_Sync " +
                    "ON UserEvents(IsSynced, Id);");
            }
        }

        private static void CreateUserEventsTable(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            string localColumns = isLocal
                ? ", DepartmentKey TEXT NOT NULL DEFAULT '" + CentralDatabasePathBuilder.UnknownDepartmentKey + "'" +
                  ", IsSynced INTEGER NOT NULL DEFAULT 0, SyncedAt TEXT NOT NULL DEFAULT ''"
                : string.Empty;

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS UserEvents (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "SyncId TEXT NOT NULL UNIQUE, " +
                "EventTime TEXT NOT NULL, " +
                "WindowsUser TEXT NOT NULL, " +
                "EventTransactionId INTEGER NOT NULL, " +
                "AddedCount INTEGER NOT NULL DEFAULT 0, " +
                "ModifiedCount INTEGER NOT NULL DEFAULT 0, " +
                "DeletedCount INTEGER NOT NULL DEFAULT 0" +
                localColumns +
                ");");
        }

        private static void EnsureAgentErrorsSchema(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            CreateAgentErrorsTable(connection, transaction, isLocal);
            if (isLocal)
            {
                EnsureColumn(
                    connection,
                    transaction,
                    "AgentErrors",
                    "DepartmentKey",
                    "TEXT NOT NULL DEFAULT '" + CentralDatabasePathBuilder.UnknownDepartmentKey + "'");
            }

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_AgentErrors_Time " +
                "ON AgentErrors(ErrorTime);");

            if (isLocal)
            {
                ExecuteNonQuery(
                    connection,
                    transaction,
                    "CREATE INDEX IF NOT EXISTS IX_AgentErrors_Sync " +
                    "ON AgentErrors(IsSynced, Id);");
            }
        }

        private static void CreateAgentErrorsTable(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            string localColumns = isLocal
                ? ", DepartmentKey TEXT NOT NULL DEFAULT '" + CentralDatabasePathBuilder.UnknownDepartmentKey + "'" +
                  ", IsSynced INTEGER NOT NULL DEFAULT 0, SyncedAt TEXT NOT NULL DEFAULT ''"
                : string.Empty;

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS AgentErrors (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "SyncId TEXT NOT NULL UNIQUE, " +
                "ErrorTime TEXT NOT NULL, " +
                "WindowsUser TEXT NOT NULL, " +
                "Source TEXT NOT NULL, " +
                "ErrorType TEXT NOT NULL, " +
                "ErrorMessage TEXT NOT NULL, " +
                "ErrorStackTrace TEXT NOT NULL" +
                localColumns +
                ");");
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

        private static void EnsureEventTransactionsSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS EventTransactions (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "EventName TEXT NOT NULL, " +
                "TransactionName TEXT NOT NULL, " +
                "UNIQUE(EventName, TransactionName)" +
                ");");
        }

        private string GetCentralEventDatabasePath(UserEventRecord userEvent)
        {
            return GetCentralDatabasePath(userEvent.EventTime, userEvent.DepartmentKey);
        }

        private string GetCentralErrorDatabasePath(AgentErrorRecord error)
        {
            return GetCentralDatabasePath(error.ErrorTime, error.DepartmentKey);
        }

        private string GetCentralDatabasePath(string recordTime, string departmentKey)
        {
            if (CentralDatabasePathBuilder.IsDatabaseFilePath(_centralDatabasePath))
                return _centralDatabasePath;

            return CentralDatabasePathBuilder.GetDatabasePath(_centralDatabasePath, departmentKey, recordTime);
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
                ExecuteNonQuery(connection, null, "PRAGMA busy_timeout=" + ModuleData.LocalBusyTimeoutMs + ";");
                ExecuteNonQuery(connection, null, "PRAGMA wal_checkpoint(TRUNCATE);");
            }
        }

        private string CreateArchiveDatabasePath()
        {
            string directoryPath = Path.GetDirectoryName(_localDatabasePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(_localDatabasePath);
            string extension = Path.GetExtension(_localDatabasePath);
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
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

        private IEnumerable<string> GetLocalQueueDatabasePaths()
        {
            List<string> paths = new List<string>();
            string directoryPath = Path.GetDirectoryName(_localDatabasePath);
            if (!string.IsNullOrWhiteSpace(directoryPath) && Directory.Exists(directoryPath))
            {
                string archivePattern = Path.GetFileNameWithoutExtension(_localDatabasePath) + "_*" + Path.GetExtension(_localDatabasePath);
                paths.AddRange(Directory.GetFiles(directoryPath, archivePattern).OrderBy(p => p, StringComparer.OrdinalIgnoreCase));
            }

            if (File.Exists(_localDatabasePath))
                paths.Add(_localDatabasePath);

            return paths;
        }

        private void CleanupSyncedArchives()
        {
            lock (_localLock)
            {
                string directoryPath = Path.GetDirectoryName(_localDatabasePath);
                if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
                    return;

                string archivePattern = Path.GetFileNameWithoutExtension(_localDatabasePath) + "_*" + Path.GetExtension(_localDatabasePath);
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
                return;

            int retentionMonths = ModuleData.CentralDatabaseRetentionMonths;
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
                    "KPLN_UserDataAgent_" + departmentKey + "_" + monthDirectoryName + ".db");
                DeleteDatabaseFiles(databasePath);
                TryDeleteDirectoryIfEmpty(fullMonthDirectoryPath);
            }
            catch
            {
            }
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

        private bool HasPendingRows(string databasePath)
        {
            using (SQLiteConnection connection = CreateConnection(databasePath, ModuleData.LocalBusyTimeoutMs))
            {
                connection.Open();
                EnsureLocalSchema(connection, null);
                return CountPendingRows(connection, "UserEvents") > 0
                    || CountPendingRows(connection, "AgentErrors") > 0;
            }
        }

        private static int CountPendingRows(SQLiteConnection connection, string tableName)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = "SELECT COUNT(*) FROM " + tableName + " WHERE IsSynced=0;";
                return Convert.ToInt32(command.ExecuteScalar());
            }
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

        private static long UpsertEventTransaction(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string tableName,
            string eventName,
            string transactionName)
        {
            string safeEventName = eventName ?? string.Empty;
            string safeTransactionName = transactionName ?? string.Empty;
            string quotedTableName = QuoteIdentifier(tableName);

            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO " + quotedTableName + " (EventName, TransactionName) " +
                    "VALUES (@EventName, @TransactionName);";
                command.Parameters.AddWithValue("@EventName", safeEventName);
                command.Parameters.AddWithValue("@TransactionName", safeTransactionName);
                command.ExecuteNonQuery();
            }

            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "SELECT Id FROM " + quotedTableName + " " +
                    "WHERE EventName=@EventName AND TransactionName=@TransactionName " +
                    "LIMIT 1;";
                command.Parameters.AddWithValue("@EventName", safeEventName);
                command.Parameters.AddWithValue("@TransactionName", safeTransactionName);
                object result = command.ExecuteScalar();
                return result == null || result == DBNull.Value ? 0L : Convert.ToInt64(result);
            }
        }

        private static void InsertLocalUserEvent(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            UserEventRecord userEvent)
        {
            userEvent.EventTransactionId = UpsertEventTransaction(
                connection,
                transaction,
                "EventTransactions",
                userEvent.EventName,
                userEvent.TransactionName);

            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO UserEvents (" +
                    "SyncId, EventTime, WindowsUser, EventTransactionId, " +
                    "AddedCount, ModifiedCount, DeletedCount, DepartmentKey, IsSynced, SyncedAt" +
                    ") VALUES (" +
                    "@SyncId, @EventTime, @WindowsUser, @EventTransactionId, " +
                    "@AddedCount, @ModifiedCount, @DeletedCount, @DepartmentKey, 0, ''" +
                    ");";
                AddUserEventParameters(command, userEvent);
                command.ExecuteNonQuery();
            }
        }

        private static void InsertCentralUserEvent(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            UserEventRecord userEvent)
        {
            userEvent.EventTransactionId = UpsertEventTransaction(
                connection,
                transaction,
                "EventTransactions",
                userEvent.EventName,
                userEvent.TransactionName);

            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO UserEvents (" +
                    "SyncId, EventTime, WindowsUser, EventTransactionId, " +
                    "AddedCount, ModifiedCount, DeletedCount" +
                    ") VALUES (" +
                    "@SyncId, @EventTime, @WindowsUser, @EventTransactionId, " +
                    "@AddedCount, @ModifiedCount, @DeletedCount" +
                    ");";
                AddUserEventParameters(command, userEvent);
                command.ExecuteNonQuery();
            }
        }

        private static void InsertLocalAgentError(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            AgentErrorRecord error)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO AgentErrors (" +
                    "SyncId, ErrorTime, WindowsUser, Source, " +
                    "ErrorType, ErrorMessage, ErrorStackTrace, DepartmentKey, IsSynced, SyncedAt" +
                    ") VALUES (" +
                    "@SyncId, @ErrorTime, @WindowsUser, @Source, " +
                    "@ErrorType, @ErrorMessage, @ErrorStackTrace, @DepartmentKey, 0, ''" +
                    ");";
                AddAgentErrorParameters(command, error);
                command.ExecuteNonQuery();
            }
        }

        private static void InsertCentralAgentError(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            AgentErrorRecord error)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO AgentErrors (" +
                    "SyncId, ErrorTime, WindowsUser, Source, " +
                    "ErrorType, ErrorMessage, ErrorStackTrace" +
                    ") VALUES (" +
                    "@SyncId, @ErrorTime, @WindowsUser, @Source, " +
                    "@ErrorType, @ErrorMessage, @ErrorStackTrace" +
                    ");";
                AddAgentErrorParameters(command, error);
                command.ExecuteNonQuery();
            }
        }

        private static void AddUserEventParameters(SQLiteCommand command, UserEventRecord userEvent)
        {
            command.Parameters.AddWithValue("@SyncId", userEvent.SyncId);
            command.Parameters.AddWithValue("@EventTime", userEvent.EventTime);
            command.Parameters.AddWithValue("@WindowsUser", userEvent.WindowsUser ?? string.Empty);
            AddDepartmentKeyParameterIfNeeded(command, userEvent.DepartmentKey);
            command.Parameters.AddWithValue("@EventTransactionId", userEvent.EventTransactionId);
            command.Parameters.AddWithValue("@AddedCount", userEvent.AddedCount);
            command.Parameters.AddWithValue("@ModifiedCount", userEvent.ModifiedCount);
            command.Parameters.AddWithValue("@DeletedCount", userEvent.DeletedCount);
        }

        private static void AddAgentErrorParameters(SQLiteCommand command, AgentErrorRecord error)
        {
            command.Parameters.AddWithValue("@SyncId", error.SyncId);
            command.Parameters.AddWithValue("@ErrorTime", error.ErrorTime);
            command.Parameters.AddWithValue("@WindowsUser", error.WindowsUser ?? string.Empty);
            AddDepartmentKeyParameterIfNeeded(command, error.DepartmentKey);
            command.Parameters.AddWithValue("@Source", error.Source ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorType", error.ErrorType ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorMessage", error.ErrorMessage ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorStackTrace", error.ErrorStackTrace ?? string.Empty);
        }

        private static void AddDepartmentKeyParameterIfNeeded(SQLiteCommand command, string departmentKey)
        {
            if (command.CommandText.IndexOf("@DepartmentKey", StringComparison.OrdinalIgnoreCase) >= 0)
                command.Parameters.AddWithValue("@DepartmentKey", CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentKey));
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

        private sealed class PendingSyncBatch
        {
            public PendingSyncBatch(UserEventRecord[] userEvents, AgentErrorRecord[] errors)
            {
                UserEvents = userEvents;
                Errors = errors;
            }

            public UserEventRecord[] UserEvents { get; private set; }
            public AgentErrorRecord[] Errors { get; private set; }

            public bool HasRows
            {
                get { return UserEvents.Length > 0 || Errors.Length > 0; }
            }
        }
    }
}