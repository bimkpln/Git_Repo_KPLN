using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;

namespace KPLN_CoordiantorAI.Common
{
    public sealed class CoordinatorAiRepository
    {
        private const int SchemaVersion = 1;

        private const string SettingsTableName = "Settings";
        private const string SessionsTableName = "QuestionSessions";
        private const string BitrixCoordinatorsTableName = "Bitrix24DepartmentCoordinators";

        public string DatabaseFolder
        {
            get { return CoordinatorAiDatabaseConfig.DatabaseFolder; }
        }

        public string DatabaseFilePath
        {
            get { return CoordinatorAiDatabaseConfig.DatabaseFilePath; }
        }

        public bool DatabaseExists
        {
            get { return File.Exists(DatabaseFilePath); }
        }

        private string ConnectionString
        {
            get { return string.Format("Data Source={0};Version=3;", DatabaseFilePath); }
        }

        public void EnsureDatabase()
        {
            if (!Directory.Exists(DatabaseFolder))
                Directory.CreateDirectory(DatabaseFolder);

            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                ExecuteNonQuery(connection, null, GetCreateSettingsTableSql());
                ExecuteNonQuery(connection, null, GetCreateSessionsTableSql());
                ExecuteNonQuery(connection, null, GetCreateBitrixCoordinatorsTableSql());
                EnsureSessionColumns(connection);
                EnsureBitrixCoordinatorColumns(connection);
                ExecuteNonQuery(connection, null, string.Format("PRAGMA user_version = {0};", SchemaVersion));
            }
        }

        public GigaChatSettings LoadGigaChatSettings()
        {
            GigaChatSettings settings = new GigaChatSettings();
            if (!DatabaseExists)
                return settings;

            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                settings.AuthUrl = GetSetting(connection, "GigaChat.AuthUrl", settings.AuthUrl);
                settings.ApiUrl = GetSetting(connection, "GigaChat.ApiUrl", settings.ApiUrl);
                settings.ClientId = GetSetting(connection, "GigaChat.ClientId", string.Empty);
                settings.ClientSecret = GetSetting(connection, "GigaChat.ClientSecret", string.Empty);
                settings.Scope = GetSetting(connection, "GigaChat.Scope", settings.Scope);
                settings.CertificatePath = GetSetting(connection, "GigaChat.CertificatePath", string.Empty);
                settings.EmbeddingFolderPath = GetEmbeddingFolderPath(connection);
                settings.SystemPrompt = GetSetting(connection, "GigaChat.SystemPrompt", settings.SystemPrompt);
                settings.ResponseContextPrompt = GetSetting(connection, "GigaChat.ResponseContextPrompt", settings.ResponseContextPrompt);
                settings.ArticleHintPrompt = GetSetting(connection, "GigaChat.ArticleHintPrompt", settings.ArticleHintPrompt);
                settings.AiSearchSettingsJson = GetSetting(connection, "GigaChat.AiSearchSettingsJson", settings.AiSearchSettingsJson);
                settings.ArticleAliasesJson = GetSetting(connection, "GigaChat.ArticleAliasesJson", settings.ArticleAliasesJson);
            }

            return settings;
        }

        public void SaveGigaChatSettings(GigaChatSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {
                    SetSetting(connection, transaction, "GigaChat.AuthUrl", settings.AuthUrl);
                    SetSetting(connection, transaction, "GigaChat.ApiUrl", settings.ApiUrl);
                    SetSetting(connection, transaction, "GigaChat.ClientId", settings.ClientId);
                    SetSetting(connection, transaction, "GigaChat.ClientSecret", settings.ClientSecret);
                    SetSetting(connection, transaction, "GigaChat.Scope", settings.Scope);
                    SetSetting(connection, transaction, "GigaChat.CertificatePath", settings.CertificatePath);
                    SetSetting(connection, transaction, "GigaChat.EmbeddingFolderPath", settings.EmbeddingFolderPath);
                    SetSetting(connection, transaction, "GigaChat.SystemPrompt", settings.SystemPrompt);
                    SetSetting(connection, transaction, "GigaChat.ResponseContextPrompt", settings.ResponseContextPrompt);
                    SetSetting(connection, transaction, "GigaChat.ArticleHintPrompt", settings.ArticleHintPrompt);
                    SetSetting(connection, transaction, "GigaChat.AiSearchSettingsJson", settings.AiSearchSettingsJson);
                    SetSetting(connection, transaction, "GigaChat.ArticleAliasesJson", settings.ArticleAliasesJson);
                    transaction.Commit();
                }
            }
        }

        public Bitrix24Settings LoadBitrix24Settings()
        {
            Bitrix24Settings settings = CreateDefaultBitrix24Settings();
            if (!DatabaseExists)
                return settings;

            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                settings.WebhookUrl = GetSetting(connection, "Bitrix24.WebhookUrl", string.Empty);
                settings.CoordinatorMessageMode = ParseBitrixMessageMode(GetSetting(connection, "Bitrix24.CoordinatorMessageMode", settings.CoordinatorMessageMode.ToString()));

                Dictionary<int, Bitrix24DepartmentCoordinator> savedCoordinators = new Dictionary<int, Bitrix24DepartmentCoordinator>();
                using (SQLiteCommand command = new SQLiteCommand(GetSelectBitrixCoordinatorsSql(), connection))
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Bitrix24DepartmentCoordinator coordinator = ReadBitrixCoordinator(reader);
                        savedCoordinators[coordinator.DepartmentId] = coordinator;
                    }
                }

                for (int i = 0; i < settings.DepartmentCoordinators.Count; i++)
                {
                    Bitrix24DepartmentCoordinator savedCoordinator;
                    if (savedCoordinators.TryGetValue(settings.DepartmentCoordinators[i].DepartmentId, out savedCoordinator))
                        settings.DepartmentCoordinators[i] = savedCoordinator;
                }
            }

            return settings;
        }

        public void SaveBitrix24Settings(Bitrix24Settings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {
                    SetSetting(connection, transaction, "Bitrix24.WebhookUrl", settings.WebhookUrl);
                    SetSetting(connection, transaction, "Bitrix24.CoordinatorMessageMode", settings.CoordinatorMessageMode.ToString());

                    foreach (Bitrix24DepartmentCoordinator coordinator in settings.DepartmentCoordinators)
                    {
                        if (coordinator == null)
                            continue;

                        SaveBitrixCoordinator(connection, transaction, coordinator);
                    }

                    transaction.Commit();
                }
            }
        }

        public void DeleteAllQuestionSessions()
        {
            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {
                    ExecuteNonQuery(connection, transaction, string.Format("DELETE FROM {0};", SessionsTableName));
                    transaction.Commit();
                }
            }
        }

        public IList<ChatSession> LoadActiveChatSessionsForUser(string userName)
        {
            EnsureDatabase();
            List<ChatSession> sessions = new List<ChatSession>();

            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteCommand command = new SQLiteCommand(GetSelectSessionsForUserSql(), connection))
                {
                    command.Parameters.AddWithValue("@UserName", userName ?? string.Empty);
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                            sessions.Add(ReadChatSession(reader));
                    }
                }
            }

            return sessions;
        }

        public void MarkChatSessionDeleted(string sessionId)
        {
            if (string.IsNullOrWhiteSpace(sessionId))
                return;

            EnsureDatabase();
            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                using (SQLiteCommand command = new SQLiteCommand(
                    string.Format("UPDATE {0} SET IsDeleted=1, DateTime=@DateTime WHERE Id=@Id;", SessionsTableName),
                    connection,
                    transaction))
                {
                    command.Parameters.AddWithValue("@Id", sessionId);
                    command.Parameters.AddWithValue("@DateTime", ToDbDate(DateTime.Now));
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
            }
        }

        public void SaveChatSession(ChatSession session)
        {
            if (session == null)
                throw new ArgumentNullException(nameof(session));

            EnsureDatabase();
            DateTime saveDateTime = DateTime.Now;
            session.UpdatedAt = saveDateTime;
            string transcript = ChatTranscriptFormatter.BuildTranscript(session.Messages);

            using (SQLiteConnection connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();
                using (SQLiteTransaction transaction = connection.BeginTransaction())
                {
                    using (SQLiteCommand command = new SQLiteCommand(GetUpsertSessionSql(), connection, transaction))
                    {
                        command.Parameters.AddWithValue("@Id", session.Id);
                        command.Parameters.AddWithValue("@DateTime", ToDbDate(saveDateTime));
                        command.Parameters.AddWithValue("@UserName", session.UserName ?? string.Empty);
                        command.Parameters.AddWithValue("@SubDepartmentId", session.SubDepartmentId);
                        command.Parameters.AddWithValue("@Chat", transcript);
                        command.Parameters.AddWithValue("@Reaction", session.Reaction.ToString());
                        command.Parameters.AddWithValue("@IsDeleted", session.IsDeleted ? 1 : 0);
                        command.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
            }
        }

        private static string ToDbDate(DateTime value)
        {
            return value.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private static void ExecuteNonQuery(SQLiteConnection connection, SQLiteTransaction transaction, string query)
        {
            using (SQLiteCommand command = new SQLiteCommand(query, connection, transaction))
            {
                command.ExecuteNonQuery();
            }
        }

        private static string GetSetting(SQLiteConnection connection, string key, string defaultValue)
        {
            using (SQLiteCommand command = new SQLiteCommand(
                string.Format("SELECT Value FROM {0} WHERE Key=@Key LIMIT 1;", SettingsTableName),
                connection))
            {
                command.Parameters.AddWithValue("@Key", key);
                object result = command.ExecuteScalar();
                return result == null || result == DBNull.Value ? defaultValue : result.ToString();
            }
        }

        private static void SetSetting(SQLiteConnection connection, SQLiteTransaction transaction, string key, string value)
        {
            using (SQLiteCommand command = new SQLiteCommand(
                string.Format("INSERT OR REPLACE INTO {0} (Key, Value, UpdatedAt) VALUES (@Key, @Value, @UpdatedAt);", SettingsTableName),
                connection,
                transaction))
            {
                command.Parameters.AddWithValue("@Key", key);
                command.Parameters.AddWithValue("@Value", value ?? string.Empty);
                command.Parameters.AddWithValue("@UpdatedAt", ToDbDate(DateTime.Now));
                command.ExecuteNonQuery();
            }
        }

        private static string GetEmbeddingFolderPath(SQLiteConnection connection)
        {
            string folderPath = GetSetting(connection, "GigaChat.EmbeddingFolderPath", string.Empty);
            if (!string.IsNullOrWhiteSpace(folderPath))
                return folderPath;

            return GetLegacyEmbeddingFolderPath(GetSetting(connection, "GigaChat.EmbeddingFilePaths", string.Empty));
        }

        private static string GetLegacyEmbeddingFolderPath(string legacyFilePaths)
        {
            if (string.IsNullOrWhiteSpace(legacyFilePaths))
                return string.Empty;

            string[] filePaths = legacyFilePaths.Replace("\r\n", "\n").Split('\n');
            foreach (string filePath in filePaths)
            {
                string normalizedPath = (filePath ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(normalizedPath))
                    continue;

                if (Directory.Exists(normalizedPath))
                    return normalizedPath;

                string folderPath = Path.GetDirectoryName(normalizedPath);
                if (!string.IsNullOrWhiteSpace(folderPath))
                    return folderPath;
            }

            return string.Empty;
        }

        private static Bitrix24Settings CreateDefaultBitrix24Settings()
        {
            Bitrix24Settings settings = new Bitrix24Settings();
            foreach (SubDepartmentInfo subDepartment in SubDepartmentNameResolver.GetKnownSubDepartments())
            {
                settings.DepartmentCoordinators.Add(new Bitrix24DepartmentCoordinator
                {
                    DepartmentId = subDepartment.Id,
                    DepartmentName = subDepartment.Name
                });
            }

            return settings;
        }

        private static void SaveBitrixCoordinator(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            Bitrix24DepartmentCoordinator coordinator)
        {
            IList<Bitrix24CoordinatorContact> contacts = coordinator.GetConfiguredContacts();
            Bitrix24CoordinatorContact user1 = contacts.Count > 0 ? contacts[0] : null;
            Bitrix24CoordinatorContact user2 = contacts.Count > 1 ? contacts[1] : null;
            Bitrix24CoordinatorContact user3 = contacts.Count > 2 ? contacts[2] : null;

            using (SQLiteCommand command = new SQLiteCommand(GetUpsertBitrixCoordinatorSql(), connection, transaction))
            {
                command.Parameters.AddWithValue("@DepartmentId", coordinator.DepartmentId);
                command.Parameters.AddWithValue("@DepartmentName", coordinator.DepartmentName ?? SubDepartmentNameResolver.GetName(coordinator.DepartmentId));
                command.Parameters.AddWithValue("@NotifyAllCoordinators", coordinator.NotifyAllCoordinators ? 1 : 0);
                command.Parameters.AddWithValue("@User1Id", user1 == null ? string.Empty : user1.UserId);
                command.Parameters.AddWithValue("@User1Name", user1 == null ? string.Empty : user1.UserName);
                command.Parameters.AddWithValue("@User2Id", user2 == null ? string.Empty : user2.UserId);
                command.Parameters.AddWithValue("@User2Name", user2 == null ? string.Empty : user2.UserName);
                command.Parameters.AddWithValue("@User3Id", user3 == null ? string.Empty : user3.UserId);
                command.Parameters.AddWithValue("@User3Name", user3 == null ? string.Empty : user3.UserName);
                command.Parameters.AddWithValue("@UpdatedAt", ToDbDate(DateTime.Now));
                command.ExecuteNonQuery();
            }
        }

        private static Bitrix24DepartmentCoordinator ReadBitrixCoordinator(SQLiteDataReader reader)
        {
            int departmentId = ReadInt(reader, "DepartmentId");
            string departmentName = ReadString(reader, "DepartmentName");
            Bitrix24DepartmentCoordinator coordinator = new Bitrix24DepartmentCoordinator
            {
                DepartmentId = departmentId,
                DepartmentName = string.IsNullOrWhiteSpace(departmentName) ? SubDepartmentNameResolver.GetName(departmentId) : departmentName,
                NotifyAllCoordinators = ReadInt(reader, "NotifyAllCoordinators") != 0
            };

            AddBitrixContact(coordinator, ReadString(reader, "User1Id"), ReadString(reader, "User1Name"));
            AddBitrixContact(coordinator, ReadString(reader, "User2Id"), ReadString(reader, "User2Name"));
            AddBitrixContact(coordinator, ReadString(reader, "User3Id"), ReadString(reader, "User3Name"));
            return coordinator;
        }

        private static void AddBitrixContact(Bitrix24DepartmentCoordinator coordinator, string userId, string userName)
        {
            if (coordinator == null || string.IsNullOrWhiteSpace(userId))
                return;

            coordinator.Coordinators.Add(new Bitrix24CoordinatorContact
            {
                DepartmentId = coordinator.DepartmentId,
                DepartmentName = coordinator.DepartmentName,
                UserId = userId.Trim(),
                UserName = string.IsNullOrWhiteSpace(userName) ? userId.Trim() : userName.Trim()
            });
        }

        private static ChatSession ReadChatSession(SQLiteDataReader reader)
        {
            DateTime dateTime = FromDbDate(ReadString(reader, "DateTime"));
            ChatSession session = new ChatSession
            {
                Id = ReadString(reader, "Id"),
                CreatedAt = dateTime,
                UpdatedAt = dateTime,
                UserName = ReadString(reader, "UserName"),
                SubDepartmentId = ReadInt(reader, "SubDepartmentId"),
                Reaction = ParseReaction(ReadString(reader, "Reaction")),
                IsDeleted = ReadInt(reader, "IsDeleted") != 0
            };

            foreach (ChatMessage message in ChatTranscriptFormatter.ParseTranscript(ReadString(reader, "Chat")))
                session.Messages.Add(message);

            return session;
        }

        private static string ReadString(SQLiteDataReader reader, string columnName)
        {
            object value = reader[columnName];
            return value == null || value == DBNull.Value ? string.Empty : value.ToString();
        }

        private static int ReadInt(SQLiteDataReader reader, string columnName)
        {
            object value = reader[columnName];
            if (value == null || value == DBNull.Value)
                return 0;

            int intValue;
            return int.TryParse(value.ToString(), out intValue) ? intValue : 0;
        }

        private static ChatReaction ParseReaction(string value)
        {
            ChatReaction reaction;
            return Enum.TryParse(value, true, out reaction) ? reaction : ChatReaction.None;
        }

        private static Bitrix24CoordinatorMessageMode ParseBitrixMessageMode(string value)
        {
            Bitrix24CoordinatorMessageMode mode;
            return Enum.TryParse(value, true, out mode) ? mode : Bitrix24CoordinatorMessageMode.FullChat;
        }

        private static DateTime FromDbDate(string value)
        {
            DateTime dateTime;
            if (DateTime.TryParseExact(
                value,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out dateTime))
                return dateTime;

            return DateTime.Now;
        }

        private static void EnsureSessionColumns(SQLiteConnection connection)
        {
            bool hasDateTime = TableHasColumn(connection, SessionsTableName, "DateTime");
            bool hasCreatedAt = TableHasColumn(connection, SessionsTableName, "CreatedAt");
            bool hasUpdatedAt = TableHasColumn(connection, SessionsTableName, "UpdatedAt");
            bool hasIsDeleted = TableHasColumn(connection, SessionsTableName, "IsDeleted");

            if (!hasDateTime)
                ExecuteNonQuery(connection, null, string.Format("ALTER TABLE {0} ADD COLUMN DateTime TEXT;", SessionsTableName));

            if (!hasIsDeleted)
                ExecuteNonQuery(connection, null, string.Format("ALTER TABLE {0} ADD COLUMN IsDeleted INTEGER NOT NULL DEFAULT 0;", SessionsTableName));

            string fallbackExpression = "COALESCE(DateTime, @Now)";
            if (hasCreatedAt && hasUpdatedAt)
                fallbackExpression = "COALESCE(DateTime, UpdatedAt, CreatedAt, @Now)";
            else if (hasUpdatedAt)
                fallbackExpression = "COALESCE(DateTime, UpdatedAt, @Now)";
            else if (hasCreatedAt)
                fallbackExpression = "COALESCE(DateTime, CreatedAt, @Now)";

            using (SQLiteCommand command = new SQLiteCommand(
                string.Format("UPDATE {0} SET DateTime={1} WHERE DateTime IS NULL OR DateTime='';", SessionsTableName, fallbackExpression),
                connection))
            {
                command.Parameters.AddWithValue("@Now", ToDbDate(DateTime.Now));
                command.ExecuteNonQuery();
            }
        }

        private static void EnsureBitrixCoordinatorColumns(SQLiteConnection connection)
        {
            if (!TableHasColumn(connection, BitrixCoordinatorsTableName, "NotifyAllCoordinators"))
            {
                ExecuteNonQuery(
                    connection,
                    null,
                    string.Format("ALTER TABLE {0} ADD COLUMN NotifyAllCoordinators INTEGER NOT NULL DEFAULT 0;", BitrixCoordinatorsTableName));
            }
        }

        private static bool TableHasColumn(SQLiteConnection connection, string tableName, string columnName)
        {
            using (SQLiteCommand command = new SQLiteCommand(string.Format("PRAGMA table_info({0});", tableName), connection))
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    string currentColumnName = reader["name"] == null ? string.Empty : reader["name"].ToString();
                    if (string.Equals(currentColumnName, columnName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }

            return false;
        }

        private static string GetCreateSettingsTableSql()
        {
            return string.Format(
                "CREATE TABLE IF NOT EXISTS {0} (" +
                "Key TEXT NOT NULL PRIMARY KEY, " +
                "Value TEXT, " +
                "UpdatedAt TEXT NOT NULL);",
                SettingsTableName);
        }

        private static string GetCreateSessionsTableSql()
        {
            return string.Format(
                "CREATE TABLE IF NOT EXISTS {0} (" +
                "Id TEXT NOT NULL PRIMARY KEY, " +
                "DateTime TEXT NOT NULL, " +
                "UserName TEXT NOT NULL, " +
                "SubDepartmentId INTEGER NOT NULL, " +
                "Chat TEXT, " +
                "Reaction TEXT NOT NULL DEFAULT 'None', " +
                "IsDeleted INTEGER NOT NULL DEFAULT 0);",
                SessionsTableName);
        }

        private static string GetCreateBitrixCoordinatorsTableSql()
        {
            return string.Format(
                "CREATE TABLE IF NOT EXISTS {0} (" +
                "DepartmentId INTEGER NOT NULL PRIMARY KEY, " +
                "DepartmentName TEXT NOT NULL, " +
                "NotifyAllCoordinators INTEGER NOT NULL DEFAULT 0, " +
                "User1Id TEXT, " +
                "User1Name TEXT, " +
                "User2Id TEXT, " +
                "User2Name TEXT, " +
                "User3Id TEXT, " +
                "User3Name TEXT, " +
                "UpdatedAt TEXT NOT NULL);",
                BitrixCoordinatorsTableName);
        }

        private static string GetUpsertSessionSql()
        {
            return string.Format(
                "INSERT OR REPLACE INTO {0} " +
                "(Id, DateTime, UserName, SubDepartmentId, Chat, Reaction, IsDeleted) " +
                "VALUES (@Id, @DateTime, @UserName, @SubDepartmentId, @Chat, @Reaction, @IsDeleted);",
                SessionsTableName);
        }

        private static string GetSelectSessionsForUserSql()
        {
            return string.Format(
                "SELECT Id, DateTime, UserName, SubDepartmentId, Chat, Reaction, IsDeleted " +
                "FROM {0} " +
                "WHERE UserName=@UserName AND IsDeleted=0 " +
                "ORDER BY DateTime DESC;",
                SessionsTableName);
        }

        private static string GetUpsertBitrixCoordinatorSql()
        {
            return string.Format(
                "INSERT OR REPLACE INTO {0} " +
                "(DepartmentId, DepartmentName, NotifyAllCoordinators, User1Id, User1Name, User2Id, User2Name, User3Id, User3Name, UpdatedAt) " +
                "VALUES (@DepartmentId, @DepartmentName, @NotifyAllCoordinators, @User1Id, @User1Name, @User2Id, @User2Name, @User3Id, @User3Name, @UpdatedAt);",
                BitrixCoordinatorsTableName);
        }

        private static string GetSelectBitrixCoordinatorsSql()
        {
            return string.Format(
                "SELECT DepartmentId, DepartmentName, NotifyAllCoordinators, User1Id, User1Name, User2Id, User2Name, User3Id, User3Name " +
                "FROM {0} ORDER BY DepartmentId;",
                BitrixCoordinatorsTableName);
        }
    }
}