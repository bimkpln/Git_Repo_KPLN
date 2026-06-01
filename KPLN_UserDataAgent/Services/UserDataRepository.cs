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
        private readonly string _localDatabasePath;
        private readonly string _centralDatabasePath;
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

                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    ExecuteNonQuery(connection, null, "PRAGMA journal_mode=WAL;");
                    ExecuteNonQuery(connection, null, "PRAGMA busy_timeout=" + ModuleData.LocalBusyTimeoutMs + ";");
                    EnsureLocalSchema(connection, null);
                }
            }
        }

        public void InsertDocumentOpened(DocumentSnapshot document)
        {
            InsertEvent(UserEventRecord.Create(
                "DocumentOpened",
                string.Empty,
                document,
                UserContextSnapshot.Current()));
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
                    return 0;
                }

                EnsureDirectory(_centralDatabasePath);

                using (SQLiteConnection centralConnection = CreateConnection(_centralDatabasePath, ModuleData.CentralBusyTimeoutMs))
                {
                    centralConnection.Open();
                    ExecuteNonQuery(centralConnection, null, "PRAGMA busy_timeout=" + ModuleData.CentralBusyTimeoutMs + ";");
                    EnsureCentralSchema(centralConnection, null);

                    using (SQLiteTransaction transaction = centralConnection.BeginTransaction())
                    {
                        foreach (UserEventRecord userEvent in pending.UserEvents)
                        {
                            UpsertEventTransaction(centralConnection, transaction, userEvent.EventName, userEvent.TransactionName);
                            InsertCentralUserEvent(centralConnection, transaction, userEvent);
                        }

                        foreach (AgentErrorRecord error in pending.Errors)
                        {
                            InsertCentralAgentError(centralConnection, transaction, error);
                        }

                        transaction.Commit();
                    }
                }

                MarkEventsSynced(pending.UserEvents);
                MarkErrorsSynced(pending.Errors);
                CleanupSyncedArchives();
                return pending.UserEvents.Length + pending.Errors.Length;
            }
            finally
            {
                Interlocked.Decrement(ref _syncRotationGuard);
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
                    "SELECT Id, SyncId, EventTime, WindowsUser, SubDepartmentId, RevitVersion, " +
                    "DocumentTitle, DocumentPath, EventName, TransactionName, " +
                    "AddedCount, ModifiedCount, DeletedCount " +
                    "FROM UserEvents " +
                    "WHERE IsSynced=0 " +
                    "ORDER BY Id " +
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
                            SubDepartmentId = Convert.ToInt32(reader["SubDepartmentId"]),
                            RevitVersion = Convert.ToInt32(reader["RevitVersion"]),
                            DocumentTitle = Convert.ToString(reader["DocumentTitle"]),
                            DocumentPath = Convert.ToString(reader["DocumentPath"]),
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
                    "SELECT Id, SyncId, ErrorTime, WindowsUser, SubDepartmentId, RevitVersion, " +
                    "Source, ErrorType, ErrorMessage, ErrorStackTrace " +
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
                            SubDepartmentId = Convert.ToInt32(reader["SubDepartmentId"]),
                            RevitVersion = Convert.ToInt32(reader["RevitVersion"]),
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
            EnsureUserEventsSchema(connection, transaction, true);
            EnsureAgentErrorsSchema(connection, transaction, true);
            DropLocalEventTransactionTables(connection, transaction);
        }

        private static void EnsureCentralSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            EnsureUserEventsSchema(connection, transaction, false);
            EnsureEventTransactionSchema(connection, transaction);
            EnsureAgentErrorsSchema(connection, transaction, false);
        }

        private static void EnsureUserEventsSchema(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS UserEvents (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "SyncId TEXT NOT NULL UNIQUE, " +
                "EventTime TEXT NOT NULL, " +
                "WindowsUser TEXT NOT NULL, " +
                "SubDepartmentId INTEGER NOT NULL, " +
                "RevitVersion INTEGER NOT NULL, " +
                "DocumentTitle TEXT NOT NULL, " +
                "DocumentPath TEXT NOT NULL, " +
                "EventName TEXT NOT NULL, " +
                "TransactionName TEXT NOT NULL, " +
                "AddedCount INTEGER NOT NULL DEFAULT 0, " +
                "ModifiedCount INTEGER NOT NULL DEFAULT 0, " +
                "DeletedCount INTEGER NOT NULL DEFAULT 0" +
                (isLocal
                    ? ", IsSynced INTEGER NOT NULL DEFAULT 0, SyncedAt TEXT NOT NULL DEFAULT ''"
                    : string.Empty) +
                ");");

            string[] columns = GetTableColumns(connection, transaction, "UserEvents");
            bool hasRang = columns.Contains("Rang", StringComparer.OrdinalIgnoreCase);
            bool hasAddedCount = columns.Contains("AddedCount", StringComparer.OrdinalIgnoreCase);
            bool hasModifiedCount = columns.Contains("ModifiedCount", StringComparer.OrdinalIgnoreCase);
            bool hasDeletedCount = columns.Contains("DeletedCount", StringComparer.OrdinalIgnoreCase);
            bool hasIsSynced = columns.Contains("IsSynced", StringComparer.OrdinalIgnoreCase);
            bool hasSyncedAt = columns.Contains("SyncedAt", StringComparer.OrdinalIgnoreCase);

            if (!hasAddedCount)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents ADD COLUMN AddedCount INTEGER NOT NULL DEFAULT 0;");

            if (!hasModifiedCount)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents ADD COLUMN ModifiedCount INTEGER NOT NULL DEFAULT 0;");

            if (!hasDeletedCount)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents ADD COLUMN DeletedCount INTEGER NOT NULL DEFAULT 0;");

            if (isLocal && !hasIsSynced)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents ADD COLUMN IsSynced INTEGER NOT NULL DEFAULT 0;");

            if (isLocal && !hasSyncedAt)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents ADD COLUMN SyncedAt TEXT NOT NULL DEFAULT '';");

            if (hasRang || (!isLocal && (hasIsSynced || hasSyncedAt)))
                RecreateUserEventsTable(connection, transaction, isLocal);

            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE INDEX IF NOT EXISTS IX_UserEvents_Time " +
                "ON UserEvents(EventTime);");

            if (isLocal)
            {
                ExecuteNonQuery(
                    connection,
                    transaction,
                    "CREATE INDEX IF NOT EXISTS IX_UserEvents_Sync " +
                    "ON UserEvents(IsSynced, Id);");
            }
        }

        private static void RecreateUserEventsTable(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            string backupName = "UserEvents_Legacy_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            ExecuteNonQuery(connection, transaction, "DROP INDEX IF EXISTS IX_UserEvents_Time;");
            ExecuteNonQuery(connection, transaction, "DROP INDEX IF EXISTS IX_UserEvents_Sync;");
            ExecuteNonQuery(connection, transaction, "ALTER TABLE UserEvents RENAME TO " + backupName + ";");
            EnsureUserEventsSchema(connection, transaction, isLocal);

            string targetColumns =
                "SyncId, EventTime, WindowsUser, SubDepartmentId, RevitVersion, " +
                "DocumentTitle, DocumentPath, EventName, TransactionName, AddedCount, ModifiedCount, DeletedCount" +
                (isLocal ? ", IsSynced, SyncedAt" : string.Empty);

            string sourceColumns =
                "SyncId, EventTime, WindowsUser, SubDepartmentId, RevitVersion, " +
                "DocumentTitle, DocumentPath, EventName, TransactionName, AddedCount, ModifiedCount, DeletedCount" +
                (isLocal ? ", IsSynced, SyncedAt" : string.Empty);

            ExecuteNonQuery(
                connection,
                transaction,
                "INSERT OR IGNORE INTO UserEvents (" + targetColumns + ") " +
                "SELECT " + sourceColumns + " FROM " + backupName + ";");
        }

        private static void EnsureAgentErrorsSchema(SQLiteConnection connection, SQLiteTransaction transaction, bool isLocal)
        {
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS AgentErrors (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "SyncId TEXT NOT NULL UNIQUE, " +
                "ErrorTime TEXT NOT NULL, " +
                "WindowsUser TEXT NOT NULL, " +
                "SubDepartmentId INTEGER NOT NULL, " +
                "RevitVersion INTEGER NOT NULL, " +
                "Source TEXT NOT NULL, " +
                "ErrorType TEXT NOT NULL, " +
                "ErrorMessage TEXT NOT NULL, " +
                "ErrorStackTrace TEXT NOT NULL" +
                (isLocal
                    ? ", IsSynced INTEGER NOT NULL DEFAULT 0, SyncedAt TEXT NOT NULL DEFAULT ''"
                    : string.Empty) +
                ");");

            string[] columns = GetTableColumns(connection, transaction, "AgentErrors");
            bool hasIsSynced = columns.Contains("IsSynced", StringComparer.OrdinalIgnoreCase);
            bool hasSyncedAt = columns.Contains("SyncedAt", StringComparer.OrdinalIgnoreCase);

            if (isLocal && !hasIsSynced)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE AgentErrors ADD COLUMN IsSynced INTEGER NOT NULL DEFAULT 0;");

            if (isLocal && !hasSyncedAt)
                ExecuteNonQuery(connection, transaction, "ALTER TABLE AgentErrors ADD COLUMN SyncedAt TEXT NOT NULL DEFAULT '';");

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

        private static void EnsureEventTransactionSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            ExecuteNonQuery(
                connection,
                transaction,
                "CREATE TABLE IF NOT EXISTS EventTransactionRanks (" +
                "Id INTEGER PRIMARY KEY AUTOINCREMENT, " +
                "EventName TEXT NOT NULL, " +
                "TransactionName TEXT NOT NULL, " +
                "Rang INTEGER NULL, " +
                "UNIQUE(EventName, TransactionName)" +
                ");");

            if (IsColumnNotNull(connection, transaction, "EventTransactionRanks", "Rang"))
            {
                string backupName = "EventTransactionRanks_Legacy_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                ExecuteNonQuery(connection, transaction, "ALTER TABLE EventTransactionRanks RENAME TO " + backupName + ";");
                EnsureEventTransactionSchema(connection, transaction);
                ExecuteNonQuery(
                    connection,
                    transaction,
                    "INSERT OR IGNORE INTO EventTransactionRanks (EventName, TransactionName, Rang) " +
                    "SELECT EventName, TransactionName, NULL FROM " + backupName + ";");
            }
        }

        private static void DropLocalEventTransactionTables(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            List<string> tableNames = new List<string>();
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "SELECT name FROM sqlite_master " +
                    "WHERE type='table' AND name LIKE 'EventTransactionRanks%';";

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        tableNames.Add(Convert.ToString(reader["name"]));
                    }
                }
            }

            foreach (string tableName in tableNames)
            {
                ExecuteNonQuery(
                    connection,
                    transaction,
                    "DROP TABLE IF EXISTS " + QuoteIdentifier(tableName) + ";");
            }
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

        private static string[] GetTableColumns(SQLiteConnection connection, SQLiteTransaction transaction, string tableName)
        {
            List<string> columns = new List<string>();
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = "PRAGMA table_info(" + tableName + ");";

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(Convert.ToString(reader["name"]));
                    }
                }
            }

            return columns.ToArray();
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

        private static bool IsColumnNotNull(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string tableName,
            string columnName)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = "PRAGMA table_info(" + tableName + ");";

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string currentColumnName = Convert.ToString(reader["name"]);
                        if (string.Equals(currentColumnName, columnName, StringComparison.OrdinalIgnoreCase))
                            return Convert.ToInt32(reader["notnull"]) != 0;
                    }
                }
            }

            return false;
        }

        private static void UpsertEventTransaction(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            string eventName,
            string transactionName)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO EventTransactionRanks (EventName, TransactionName, Rang) " +
                    "VALUES (@EventName, @TransactionName, NULL);";
                command.Parameters.AddWithValue("@EventName", eventName ?? string.Empty);
                command.Parameters.AddWithValue("@TransactionName", transactionName ?? string.Empty);
                command.ExecuteNonQuery();
            }
        }

        private static void InsertLocalUserEvent(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            UserEventRecord userEvent)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO UserEvents (" +
                    "SyncId, EventTime, WindowsUser, SubDepartmentId, RevitVersion, DocumentTitle, DocumentPath, " +
                    "EventName, TransactionName, AddedCount, ModifiedCount, DeletedCount, IsSynced, SyncedAt" +
                    ") VALUES (" +
                    "@SyncId, @EventTime, @WindowsUser, @SubDepartmentId, @RevitVersion, @DocumentTitle, @DocumentPath, " +
                    "@EventName, @TransactionName, @AddedCount, @ModifiedCount, @DeletedCount, 0, ''" +
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
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText =
                    "INSERT OR IGNORE INTO UserEvents (" +
                    "SyncId, EventTime, WindowsUser, SubDepartmentId, RevitVersion, DocumentTitle, DocumentPath, " +
                    "EventName, TransactionName, AddedCount, ModifiedCount, DeletedCount" +
                    ") VALUES (" +
                    "@SyncId, @EventTime, @WindowsUser, @SubDepartmentId, @RevitVersion, @DocumentTitle, @DocumentPath, " +
                    "@EventName, @TransactionName, @AddedCount, @ModifiedCount, @DeletedCount" +
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
                    "SyncId, ErrorTime, WindowsUser, SubDepartmentId, RevitVersion, Source, " +
                    "ErrorType, ErrorMessage, ErrorStackTrace, IsSynced, SyncedAt" +
                    ") VALUES (" +
                    "@SyncId, @ErrorTime, @WindowsUser, @SubDepartmentId, @RevitVersion, @Source, " +
                    "@ErrorType, @ErrorMessage, @ErrorStackTrace, 0, ''" +
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
                    "SyncId, ErrorTime, WindowsUser, SubDepartmentId, RevitVersion, Source, " +
                    "ErrorType, ErrorMessage, ErrorStackTrace" +
                    ") VALUES (" +
                    "@SyncId, @ErrorTime, @WindowsUser, @SubDepartmentId, @RevitVersion, @Source, " +
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
            command.Parameters.AddWithValue("@SubDepartmentId", userEvent.SubDepartmentId);
            command.Parameters.AddWithValue("@RevitVersion", userEvent.RevitVersion);
            command.Parameters.AddWithValue("@DocumentTitle", userEvent.DocumentTitle ?? string.Empty);
            command.Parameters.AddWithValue("@DocumentPath", userEvent.DocumentPath ?? string.Empty);
            command.Parameters.AddWithValue("@EventName", userEvent.EventName ?? string.Empty);
            command.Parameters.AddWithValue("@TransactionName", userEvent.TransactionName ?? string.Empty);
            command.Parameters.AddWithValue("@AddedCount", userEvent.AddedCount);
            command.Parameters.AddWithValue("@ModifiedCount", userEvent.ModifiedCount);
            command.Parameters.AddWithValue("@DeletedCount", userEvent.DeletedCount);
        }

        private static void AddAgentErrorParameters(SQLiteCommand command, AgentErrorRecord error)
        {
            command.Parameters.AddWithValue("@SyncId", error.SyncId);
            command.Parameters.AddWithValue("@ErrorTime", error.ErrorTime);
            command.Parameters.AddWithValue("@WindowsUser", error.WindowsUser ?? string.Empty);
            command.Parameters.AddWithValue("@SubDepartmentId", error.SubDepartmentId);
            command.Parameters.AddWithValue("@RevitVersion", error.RevitVersion);
            command.Parameters.AddWithValue("@Source", error.Source ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorType", error.ErrorType ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorMessage", error.ErrorMessage ?? string.Empty);
            command.Parameters.AddWithValue("@ErrorStackTrace", error.ErrorStackTrace ?? string.Empty);
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