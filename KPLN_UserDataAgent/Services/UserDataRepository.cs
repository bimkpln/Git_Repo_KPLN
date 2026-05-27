using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class UserDataRepository
    {
        private readonly object _localLock = new object();
        private readonly string _localDatabasePath;
        private readonly string _centralDatabasePath;

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
                        UpsertEventTransaction(connection, transaction, userEvent.EventName, userEvent.TransactionName);
                        InsertLocalUserEvent(connection, transaction, userEvent);
                        transaction.Commit();
                    }
                }
            }
        }

        public int TrySyncPendingToCentral(int batchSize)
        {
            UserEventRecord[] pendingEvents = LoadPendingEvents(batchSize).ToArray();
            if (pendingEvents.Length == 0)
                return 0;

            EnsureDirectory(_centralDatabasePath);

            using (SQLiteConnection centralConnection = CreateConnection(_centralDatabasePath, ModuleData.CentralBusyTimeoutMs))
            {
                centralConnection.Open();
                ExecuteNonQuery(centralConnection, null, "PRAGMA busy_timeout=" + ModuleData.CentralBusyTimeoutMs + ";");
                EnsureCentralSchema(centralConnection, null);

                using (SQLiteTransaction transaction = centralConnection.BeginTransaction())
                {
                    foreach (UserEventRecord userEvent in pendingEvents)
                    {
                        UpsertEventTransaction(centralConnection, transaction, userEvent.EventName, userEvent.TransactionName);
                        InsertCentralUserEvent(centralConnection, transaction, userEvent);
                    }

                    transaction.Commit();
                }
            }

            MarkEventsSynced(pendingEvents.Select(e => e.LocalId).ToArray());
            return pendingEvents.Length;
        }

        private IEnumerable<UserEventRecord> LoadPendingEvents(int batchSize)
        {
            InitializeLocal();

            lock (_localLock)
            {
                List<UserEventRecord> result = new List<UserEventRecord>();
                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                using (SQLiteCommand command = connection.CreateCommand())
                {
                    connection.Open();
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
        }

        private void MarkEventsSynced(long[] localIds)
        {
            if (localIds == null || localIds.Length == 0)
                return;

            lock (_localLock)
            {
                using (SQLiteConnection connection = CreateConnection(_localDatabasePath, ModuleData.LocalBusyTimeoutMs))
                {
                    connection.Open();
                    using (SQLiteTransaction transaction = connection.BeginTransaction())
                    {
                        foreach (long localId in localIds)
                        {
                            using (SQLiteCommand command = connection.CreateCommand())
                            {
                                command.Transaction = transaction;
                                command.CommandText =
                                    "UPDATE UserEvents " +
                                    "SET IsSynced=1, SyncedAt=@SyncedAt " +
                                    "WHERE Id=@Id;";
                                command.Parameters.AddWithValue("@SyncedAt", DateTime.Now.ToString("yyyy.MM.dd. HH:mm:ss"));
                                command.Parameters.AddWithValue("@Id", localId);
                                command.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                }
            }
        }

        private static void EnsureLocalSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            EnsureUserEventsSchema(connection, transaction, true);
            EnsureEventTransactionSchema(connection, transaction);
        }

        private static void EnsureCentralSchema(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            EnsureUserEventsSchema(connection, transaction, false);
            EnsureEventTransactionSchema(connection, transaction);
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

        private static SQLiteConnection CreateConnection(string databasePath, int busyTimeoutMs)
        {
            return new SQLiteConnection(
                "Data Source=" + databasePath + ";Version=3;BusyTimeout=" + busyTimeoutMs + ";");
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
    }
}