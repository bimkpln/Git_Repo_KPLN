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
                EnsureSessionColumns(connection);
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
                settings.EmbeddingFilePaths = GetSetting(connection, "GigaChat.EmbeddingFilePaths", string.Empty);
                settings.SystemPrompt = GetSetting(connection, "GigaChat.SystemPrompt", settings.SystemPrompt);
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
                    SetSetting(connection, transaction, "GigaChat.EmbeddingFilePaths", settings.EmbeddingFilePaths);
                    SetSetting(connection, transaction, "GigaChat.SystemPrompt", settings.SystemPrompt);
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
    }
}