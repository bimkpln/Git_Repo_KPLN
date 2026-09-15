using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using KPLN_CoordiantorAI.ExternalAIModel.Mcp;

namespace KPLN_CoordiantorAI.ExternalModel
{
    internal sealed class ModelChatSessionRecord
    {
        public string SessionId { get; set; }
        public string UserName { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime? EndedAt { get; set; }
        public int RevitVersion { get; set; }
        public int RevitProcessId { get; set; }
        public string PluginVersion { get; set; }
        public string ModelName { get; set; }
        public string ConnectionType { get; set; }
    }

    internal sealed class ModelChatRequestRecord
    {
        public string RequestId { get; set; }
        public string SessionId { get; set; }
        public DateTime RequestTime { get; set; }
        public DateTime? ResponseTime { get; set; }
        public string UserQuestion { get; set; }
        public string FinalAnswer { get; set; }
        public string Status { get; set; }
        public string Scenario { get; set; }
        public string AiModelName { get; set; }
        public int RevitVersion { get; set; }
        public int RevitProcessId { get; set; }
        public string ModelName { get; set; }
        public string ViewName { get; set; }
        public string ModelIdentity { get; set; }
        public long? DurationMs { get; set; }
        public string UsedToolNamesJson { get; set; }
        public long? CacheHitTokens { get; set; }
        public long? CacheMissTokens { get; set; }
        public long? CompletionTokens { get; set; }
        public decimal? Cost { get; set; }
        public string ErrorMessage { get; set; }
    }

    internal sealed class ModelChatAttachmentRecord
    {
        public string Id { get; set; }
        public string RequestId { get; set; }
        public string AttachmentType { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public string ContentType { get; set; }
        public long SizeBytes { get; set; }
        public int? Width { get; set; }
        public int? Height { get; set; }
    }

    internal sealed class ModelChatHistoryEntry
    {
        public string RequestId { get; set; }
        public string SessionId { get; set; }
        public string RequestTime { get; set; }
        public string UserQuestion { get; set; }
        public string FinalAnswer { get; set; }
        public string ModelName { get; set; }
        public string ModelIdentity { get; set; }
        public int RevitVersion { get; set; }
    }

    internal sealed class ModelAnswerFeedbackRecord
    {
        public string FeedbackId { get; set; }
        public string RequestId { get; set; }
        public string Rating { get; set; }
        public string ReasonCode { get; set; }
        public string Comment { get; set; }
    }

    internal sealed class ModelChatHistoryRepository
    {
        private sealed class MissingCostRecord
        {
            public string RequestId { get; set; }
            public string RequestTime { get; set; }
            public string AiModelName { get; set; }
            public long CacheHitTokens { get; set; }
            public long CacheMissTokens { get; set; }
            public long CompletionTokens { get; set; }
        }

        private const string DatabaseFolder =
            @"Z:\Отдел BIM\Тарчоков Мухамед\2. Плагины\1. Собственные плагины\0. Shared Plugins\Logs\BimHelper\1. Логи в форме БД";
        private static readonly object SyncRoot = new object();
        private static readonly TimeZoneInfo MoscowTimeZone = ResolveMoscowTimeZone();
        private readonly string _databaseFilePath;

        public ModelChatHistoryRepository()
        {
            _databaseFilePath = Path.Combine(
                DatabaseFolder,
                "CoordinatorAI_History_" + SanitizeFileName(Environment.UserName) + ".db");
        }

        public string DatabaseFilePath
        {
            get { return _databaseFilePath; }
        }

        private string ConnectionString
        {
            get { return string.Format("Data Source={0};Version=3;", _databaseFilePath); }
        }

        public void EnsureDatabaseAndSynchronizeTools(IReadOnlyList<RevitMcpToolDefinition> tools)
        {
            lock (SyncRoot)
            {
                EnsureFolder();
                using (SQLiteConnection connection = OpenConnection())
                {
                    ExecuteNonQuery(connection, null, GetCreateSessionsSql());
                    ExecuteNonQuery(connection, null, GetCreateRequestsSql());
                    EnsureModelIdentityColumn(connection);
                    ExecuteNonQuery(connection, null, GetCreateAttachmentsSql());
                    ExecuteNonQuery(connection, null, GetCreateToolDefinitionsSql());
                    ExecuteNonQuery(connection, null, GetCreateFeedbackSql());
                    ExecuteNonQuery(connection, null, "CREATE INDEX IF NOT EXISTS IX_ModelChatRequests_SessionId ON ModelChatRequests(SessionId);");
                    ExecuteNonQuery(connection, null, "CREATE INDEX IF NOT EXISTS IX_ModelChatRequests_ModelIdentity ON ModelChatRequests(ModelIdentity, RevitVersion);");
                    ExecuteNonQuery(connection, null, "CREATE INDEX IF NOT EXISTS IX_ModelChatRequests_RequestTime ON ModelChatRequests(RequestTime);");
                    ExecuteNonQuery(connection, null, "CREATE INDEX IF NOT EXISTS IX_ModelAttachments_RequestId ON ModelAttachments(RequestId);");
                    SynchronizeToolDefinitions(connection, tools);
                    NormalizeStoredDateTimes(connection);
                    BackfillMissingCosts(connection);
                }
            }
        }

        public void SaveSession(ModelChatSessionRecord record)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.SessionId))
                throw new ArgumentException("Model chat session record is missing its SessionId.");

            lock (SyncRoot)
            using (SQLiteConnection connection = OpenConnection())
            {
                const string updateSql =
                    "UPDATE ModelChatSessions SET UserName=@UserName, StartedAt=@StartedAt, EndedAt=@EndedAt, " +
                    "RevitVersion=@RevitVersion, RevitProcessId=@RevitProcessId, PluginVersion=@PluginVersion, " +
                    "ModelName=@ModelName, ConnectionType=@ConnectionType WHERE SessionId=@SessionId;";
                using (SQLiteCommand command = CreateSessionCommand(updateSql, connection, record))
                {
                    if (command.ExecuteNonQuery() > 0)
                        return;
                }

                const string insertSql =
                    "INSERT INTO ModelChatSessions (SessionId, UserName, StartedAt, EndedAt, RevitVersion, " +
                    "RevitProcessId, PluginVersion, ModelName, ConnectionType) VALUES (@SessionId, @UserName, " +
                    "@StartedAt, @EndedAt, @RevitVersion, @RevitProcessId, @PluginVersion, @ModelName, @ConnectionType);";
                using (SQLiteCommand command = CreateSessionCommand(insertSql, connection, record))
                    command.ExecuteNonQuery();
            }
        }

        public void SaveRequest(ModelChatRequestRecord record, IList<ModelChatAttachmentRecord> attachments)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.RequestId))
                throw new ArgumentException("Model chat request record is missing its RequestId.");

            lock (SyncRoot)
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteTransaction transaction = connection.BeginTransaction())
            {
                const string updateSql =
                    "UPDATE ModelChatRequests SET SessionId=@SessionId, RequestTime=@RequestTime, ResponseTime=@ResponseTime, " +
                    "UserQuestion=@UserQuestion, FinalAnswer=@FinalAnswer, Status=@Status, Scenario=@Scenario, " +
                    "AiModelName=@AiModelName, RevitVersion=@RevitVersion, RevitProcessId=@RevitProcessId, " +
                    "ModelName=@ModelName, ModelIdentity=@ModelIdentity, ViewName=@ViewName, DurationMs=@DurationMs, UsedToolNamesJson=@UsedToolNamesJson, " +
                    "CacheHitTokens=@CacheHitTokens, CacheMissTokens=@CacheMissTokens, CompletionTokens=@CompletionTokens, " +
                    "Cost=@Cost, ErrorMessage=@ErrorMessage WHERE RequestId=@RequestId;";

                int updated;
                using (SQLiteCommand command = CreateRequestCommand(updateSql, connection, transaction, record))
                    updated = command.ExecuteNonQuery();

                if (updated == 0)
                {
                    const string insertSql =
                        "INSERT INTO ModelChatRequests (RequestId, SessionId, RequestTime, ResponseTime, UserQuestion, " +
                        "FinalAnswer, Status, Scenario, AiModelName, RevitVersion, RevitProcessId, ModelName, ModelIdentity, ViewName, " +
                        "DurationMs, UsedToolNamesJson, CacheHitTokens, CacheMissTokens, CompletionTokens, Cost, ErrorMessage) " +
                        "VALUES (@RequestId, @SessionId, @RequestTime, @ResponseTime, @UserQuestion, @FinalAnswer, @Status, " +
                        "@Scenario, @AiModelName, @RevitVersion, @RevitProcessId, @ModelName, @ModelIdentity, @ViewName, @DurationMs, " +
                        "@UsedToolNamesJson, @CacheHitTokens, @CacheMissTokens, @CompletionTokens, @Cost, @ErrorMessage);";
                    using (SQLiteCommand command = CreateRequestCommand(insertSql, connection, transaction, record))
                        command.ExecuteNonQuery();
                }

                using (SQLiteCommand deleteAttachments = new SQLiteCommand(
                    "DELETE FROM ModelAttachments WHERE RequestId=@RequestId;", connection, transaction))
                {
                    deleteAttachments.Parameters.AddWithValue("@RequestId", record.RequestId);
                    deleteAttachments.ExecuteNonQuery();
                }

                if (attachments != null)
                {
                    foreach (ModelChatAttachmentRecord attachment in attachments)
                        InsertAttachment(connection, transaction, attachment);
                }

                transaction.Commit();
            }
        }

        public IList<ModelChatHistoryEntry> GetCompletedHistoryPage(int offset, int pageSize, out bool hasMore)
        {
            int safeOffset = Math.Max(0, offset);
            int safePageSize = Math.Max(1, Math.Min(100, pageSize));
            List<ModelChatHistoryEntry> result = new List<ModelChatHistoryEntry>();

            lock (SyncRoot)
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = new SQLiteCommand(
                "SELECT RequestId, SessionId, RequestTime, UserQuestion, FinalAnswer, ModelName, RevitVersion, " +
                "CASE WHEN ModelIdentity IS NULL OR TRIM(ModelIdentity)='' " +
                "THEN 'legacy-session:' || SessionId ELSE ModelIdentity END AS HistoryModelIdentity " +
                "FROM ModelChatRequests WHERE Scenario='wpf_window' " +
                "AND FinalAnswer IS NOT NULL AND TRIM(FinalAnswer)<>'' " +
                "ORDER BY RequestTime DESC LIMIT @Limit OFFSET @Offset;", connection))
            {
                command.Parameters.AddWithValue("@Limit", safePageSize + 1);
                command.Parameters.AddWithValue("@Offset", safeOffset);
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new ModelChatHistoryEntry
                        {
                            RequestId = reader["RequestId"].ToString(),
                            SessionId = reader["SessionId"].ToString(),
                            RequestTime = reader["RequestTime"].ToString(),
                            UserQuestion = reader["UserQuestion"].ToString(),
                            FinalAnswer = reader["FinalAnswer"].ToString(),
                            ModelName = reader["ModelName"].ToString(),
                            ModelIdentity = reader["HistoryModelIdentity"].ToString(),
                            RevitVersion = Convert.ToInt32(reader["RevitVersion"], CultureInfo.InvariantCulture)
                        });
                    }
                }
            }

            hasMore = result.Count > safePageSize;
            if (hasMore)
                result.RemoveAt(result.Count - 1);
            return result;
        }

        public void SaveFeedback(ModelAnswerFeedbackRecord record)
        {
            if (record == null || string.IsNullOrWhiteSpace(record.RequestId))
                throw new ArgumentException("Model answer feedback is missing its RequestId.");
            if (!string.Equals(record.Rating, "Positive", StringComparison.Ordinal)
                && !string.Equals(record.Rating, "Negative", StringComparison.Ordinal))
                throw new ArgumentException("Model answer feedback has an unsupported rating.");

            lock (SyncRoot)
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteTransaction transaction = connection.BeginTransaction())
            {
                const string updateSql =
                    "UPDATE ModelAnswerFeedback SET Rating=@Rating, ReasonCode=@ReasonCode, Comment=@Comment " +
                    "WHERE RequestId=@RequestId;";
                int updated;
                using (SQLiteCommand command = CreateFeedbackCommand(updateSql, connection, transaction, record))
                    updated = command.ExecuteNonQuery();

                if (updated == 0)
                {
                    const string insertSql =
                        "INSERT INTO ModelAnswerFeedback (FeedbackId, RequestId, Rating, ReasonCode, Comment) " +
                        "VALUES (@FeedbackId, @RequestId, @Rating, @ReasonCode, @Comment);";
                    using (SQLiteCommand command = CreateFeedbackCommand(insertSql, connection, transaction, record))
                        command.ExecuteNonQuery();
                }

                transaction.Commit();
            }
        }

        public void DeleteFeedback(string requestId)
        {
            if (string.IsNullOrWhiteSpace(requestId))
                return;

            lock (SyncRoot)
            using (SQLiteConnection connection = OpenConnection())
            using (SQLiteCommand command = new SQLiteCommand(
                "DELETE FROM ModelAnswerFeedback WHERE RequestId=@RequestId;", connection))
            {
                command.Parameters.AddWithValue("@RequestId", requestId);
                command.ExecuteNonQuery();
            }
        }

        private static void EnsureModelIdentityColumn(SQLiteConnection connection)
        {
            bool exists = false;
            using (SQLiteCommand command = new SQLiteCommand("PRAGMA table_info(ModelChatRequests);", connection))
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (string.Equals(reader["name"].ToString(), "ModelIdentity", StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }
            }

            if (!exists)
            {
                ExecuteNonQuery(
                    connection,
                    null,
                    "ALTER TABLE ModelChatRequests ADD COLUMN ModelIdentity TEXT;");
            }
        }

        private SQLiteConnection OpenConnection()
        {
            EnsureFolder();
            SQLiteConnection connection = new SQLiteConnection(ConnectionString);
            connection.Open();
            using (SQLiteCommand command = new SQLiteCommand("PRAGMA busy_timeout=5000;", connection))
                command.ExecuteNonQuery();
            return connection;
        }

        private static void SynchronizeToolDefinitions(
            SQLiteConnection connection,
            IReadOnlyList<RevitMcpToolDefinition> tools)
        {
            if (tools == null)
                return;

            using (SQLiteTransaction transaction = connection.BeginTransaction())
            {
                int nextToolId = GetNextToolId(connection, transaction);
                foreach (RevitMcpToolDefinition tool in tools)
                {
                    if (tool == null || string.IsNullOrWhiteSpace(tool.Name))
                        continue;

                    int? existingToolId = GetToolId(connection, transaction, tool.Name);
                    int toolId = existingToolId ?? nextToolId++;
                    string paginationUnit;
                    int? defaultPageSize;
                    bool supportsPagination = GetPaginationMetadata(tool, out paginationUnit, out defaultPageSize);

                    const string sql =
                        "INSERT OR REPLACE INTO ToolDefinitions (ToolId, ToolName, Description, OperationType, IsReadOnly, " +
                        "SupportsPagination, PaginationUnit, DefaultPageSize) VALUES (@ToolId, @ToolName, @Description, " +
                        "@OperationType, @IsReadOnly, @SupportsPagination, @PaginationUnit, @DefaultPageSize);";
                    using (SQLiteCommand command = new SQLiteCommand(sql, connection, transaction))
                    {
                        command.Parameters.AddWithValue("@ToolId", toolId);
                        command.Parameters.AddWithValue("@ToolName", tool.Name);
                        command.Parameters.AddWithValue("@Description", tool.Description ?? string.Empty);
                        command.Parameters.AddWithValue("@OperationType", tool.Area ?? string.Empty);
                        command.Parameters.AddWithValue("@IsReadOnly", tool.RiskLevel == RevitMcpToolRiskLevel.ReadOnly ? 1 : 0);
                        command.Parameters.AddWithValue("@SupportsPagination", supportsPagination ? 1 : 0);
                        command.Parameters.AddWithValue("@PaginationUnit", DbValue(paginationUnit));
                        command.Parameters.AddWithValue("@DefaultPageSize", DbValue(defaultPageSize));
                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
        }

        private static bool GetPaginationMetadata(
            RevitMcpToolDefinition tool,
            out string paginationUnit,
            out int? defaultPageSize)
        {
            paginationUnit = null;
            defaultPageSize = null;

            if (string.Equals(tool.Name, "search_attached_file", StringComparison.Ordinal))
            {
                paginationUnit = "Matches";
                defaultPageSize = 20;
                return true;
            }

            Newtonsoft.Json.Linq.JObject properties = tool.InputSchema == null
                ? null
                : tool.InputSchema["properties"] as Newtonsoft.Json.Linq.JObject;
            if (properties == null || properties["limit"] == null || properties["offset"] == null)
                return false;

            if (string.Equals(tool.Name, "get_journal_entries_since", StringComparison.Ordinal)
                || string.Equals(tool.Name, "read_attached_file", StringComparison.Ordinal))
            {
                paginationUnit = "Characters";
                defaultPageSize = 204800;
            }
            else if (string.Equals(tool.Name, "get_viewports_and_schedules_on_sheets", StringComparison.Ordinal))
            {
                paginationUnit = "Sheets";
                defaultPageSize = 1;
            }
            else
            {
                paginationUnit = "Elements";
                defaultPageSize = 200;
            }

            return true;
        }

        private static void BackfillMissingCosts(SQLiteConnection connection)
        {
            List<MissingCostRecord> records = new List<MissingCostRecord>();
            const string selectSql =
                "SELECT RequestId, RequestTime, AiModelName, CacheHitTokens, CacheMissTokens, CompletionTokens " +
                "FROM ModelChatRequests WHERE Cost IS NULL AND CacheHitTokens IS NOT NULL " +
                "AND CacheMissTokens IS NOT NULL AND CompletionTokens IS NOT NULL;";
            using (SQLiteCommand command = new SQLiteCommand(selectSql, connection))
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    records.Add(new MissingCostRecord
                    {
                        RequestId = reader["RequestId"].ToString(),
                        RequestTime = reader["RequestTime"].ToString(),
                        AiModelName = reader["AiModelName"].ToString(),
                        CacheHitTokens = Convert.ToInt64(reader["CacheHitTokens"], CultureInfo.InvariantCulture),
                        CacheMissTokens = Convert.ToInt64(reader["CacheMissTokens"], CultureInfo.InvariantCulture),
                        CompletionTokens = Convert.ToInt64(reader["CompletionTokens"], CultureInfo.InvariantCulture)
                    });
                }
            }

            using (SQLiteTransaction transaction = connection.BeginTransaction())
            {
                foreach (MissingCostRecord record in records)
                {
                    DateTime requestTimeUtc;
                    if (!TryParseDbDateToUtc(record.RequestTime, out requestTimeUtc))
                        continue;

                    decimal cost;
                    if (!DeepSeekUsageCostCalculator.TryCalculateUsd(
                        record.AiModelName,
                        requestTimeUtc,
                        record.CacheHitTokens,
                        record.CacheMissTokens,
                        record.CompletionTokens,
                        out cost))
                        continue;

                    using (SQLiteCommand command = new SQLiteCommand(
                        "UPDATE ModelChatRequests SET Cost=@Cost WHERE RequestId=@RequestId AND Cost IS NULL;",
                        connection,
                        transaction))
                    {
                        command.Parameters.AddWithValue("@Cost", cost);
                        command.Parameters.AddWithValue("@RequestId", record.RequestId);
                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
        }

        private static void NormalizeStoredDateTimes(SQLiteConnection connection)
        {
            NormalizeDateColumn(connection, "ModelChatSessions", "SessionId", "StartedAt");
            NormalizeDateColumn(connection, "ModelChatSessions", "SessionId", "EndedAt");
            NormalizeDateColumn(connection, "ModelChatRequests", "RequestId", "RequestTime");
            NormalizeDateColumn(connection, "ModelChatRequests", "RequestId", "ResponseTime");
        }

        private static void NormalizeDateColumn(
            SQLiteConnection connection,
            string tableName,
            string idColumn,
            string dateColumn)
        {
            List<KeyValuePair<string, string>> updates = new List<KeyValuePair<string, string>>();
            string selectSql = string.Format(
                "SELECT {0}, {1} FROM {2} WHERE {1} IS NOT NULL AND {1} <> '';",
                idColumn,
                dateColumn,
                tableName);
            using (SQLiteCommand command = new SQLiteCommand(selectSql, connection))
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    string id = reader[idColumn].ToString();
                    string storedValue = reader[dateColumn].ToString();
                    DateTime utc;
                    if (!TryParseDbDateToUtc(storedValue, out utc))
                        continue;

                    string normalizedValue = FormatMoscowDate(utc);
                    if (!string.Equals(storedValue, normalizedValue, StringComparison.Ordinal))
                        updates.Add(new KeyValuePair<string, string>(id, normalizedValue));
                }
            }

            if (updates.Count == 0)
                return;

            using (SQLiteTransaction transaction = connection.BeginTransaction())
            {
                string updateSql = string.Format(
                    "UPDATE {0} SET {1}=@DateValue WHERE {2}=@Id;",
                    tableName,
                    dateColumn,
                    idColumn);
                foreach (KeyValuePair<string, string> update in updates)
                {
                    using (SQLiteCommand command = new SQLiteCommand(updateSql, connection, transaction))
                    {
                        command.Parameters.AddWithValue("@DateValue", update.Value);
                        command.Parameters.AddWithValue("@Id", update.Key);
                        command.ExecuteNonQuery();
                    }
                }

                transaction.Commit();
            }
        }

        private static int? GetToolId(SQLiteConnection connection, SQLiteTransaction transaction, string toolName)
        {
            using (SQLiteCommand command = new SQLiteCommand(
                "SELECT ToolId FROM ToolDefinitions WHERE ToolName=@ToolName LIMIT 1;", connection, transaction))
            {
                command.Parameters.AddWithValue("@ToolName", toolName);
                object value = command.ExecuteScalar();
                return value == null || value == DBNull.Value
                    ? (int?)null
                    : Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
        }

        private static int GetNextToolId(SQLiteConnection connection, SQLiteTransaction transaction)
        {
            using (SQLiteCommand command = new SQLiteCommand(
                "SELECT COALESCE(MAX(ToolId), 0) + 1 FROM ToolDefinitions;", connection, transaction))
            {
                return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
            }
        }

        private static SQLiteCommand CreateSessionCommand(
            string sql,
            SQLiteConnection connection,
            ModelChatSessionRecord record)
        {
            SQLiteCommand command = new SQLiteCommand(sql, connection);
            command.Parameters.AddWithValue("@SessionId", record.SessionId);
            command.Parameters.AddWithValue("@UserName", record.UserName ?? string.Empty);
            command.Parameters.AddWithValue("@StartedAt", ToDbDate(record.StartedAt));
            command.Parameters.AddWithValue("@EndedAt", DbDateValue(record.EndedAt));
            command.Parameters.AddWithValue("@RevitVersion", record.RevitVersion);
            command.Parameters.AddWithValue("@RevitProcessId", record.RevitProcessId);
            command.Parameters.AddWithValue("@PluginVersion", record.PluginVersion ?? string.Empty);
            command.Parameters.AddWithValue("@ModelName", record.ModelName ?? string.Empty);
            command.Parameters.AddWithValue("@ConnectionType", record.ConnectionType ?? string.Empty);
            return command;
        }

        private static SQLiteCommand CreateRequestCommand(
            string sql,
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            ModelChatRequestRecord record)
        {
            SQLiteCommand command = new SQLiteCommand(sql, connection, transaction);
            command.Parameters.AddWithValue("@RequestId", record.RequestId);
            command.Parameters.AddWithValue("@SessionId", record.SessionId ?? string.Empty);
            command.Parameters.AddWithValue("@RequestTime", ToDbDate(record.RequestTime));
            command.Parameters.AddWithValue("@ResponseTime", DbDateValue(record.ResponseTime));
            command.Parameters.AddWithValue("@UserQuestion", record.UserQuestion ?? string.Empty);
            command.Parameters.AddWithValue("@FinalAnswer", DbValue(record.FinalAnswer));
            command.Parameters.AddWithValue("@Status", record.Status ?? string.Empty);
            command.Parameters.AddWithValue("@Scenario", record.Scenario ?? string.Empty);
            command.Parameters.AddWithValue("@AiModelName", record.AiModelName ?? string.Empty);
            command.Parameters.AddWithValue("@RevitVersion", record.RevitVersion);
            command.Parameters.AddWithValue("@RevitProcessId", record.RevitProcessId);
            command.Parameters.AddWithValue("@ModelName", record.ModelName ?? string.Empty);
            command.Parameters.AddWithValue("@ModelIdentity", DbValue(record.ModelIdentity));
            command.Parameters.AddWithValue("@ViewName", record.ViewName ?? string.Empty);
            command.Parameters.AddWithValue("@DurationMs", DbValue(record.DurationMs));
            command.Parameters.AddWithValue("@UsedToolNamesJson", record.UsedToolNamesJson ?? "[]");
            command.Parameters.AddWithValue("@CacheHitTokens", DbValue(record.CacheHitTokens));
            command.Parameters.AddWithValue("@CacheMissTokens", DbValue(record.CacheMissTokens));
            command.Parameters.AddWithValue("@CompletionTokens", DbValue(record.CompletionTokens));
            command.Parameters.AddWithValue("@Cost", DbValue(record.Cost));
            command.Parameters.AddWithValue("@ErrorMessage", DbValue(record.ErrorMessage));
            return command;
        }

        private static SQLiteCommand CreateFeedbackCommand(
            string sql,
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            ModelAnswerFeedbackRecord record)
        {
            SQLiteCommand command = new SQLiteCommand(sql, connection, transaction);
            command.Parameters.AddWithValue("@FeedbackId", string.IsNullOrWhiteSpace(record.FeedbackId)
                ? Guid.NewGuid().ToString("N")
                : record.FeedbackId);
            command.Parameters.AddWithValue("@RequestId", record.RequestId);
            command.Parameters.AddWithValue("@Rating", record.Rating);
            command.Parameters.AddWithValue("@ReasonCode", DbValue(record.ReasonCode));
            command.Parameters.AddWithValue("@Comment", DbValue(record.Comment));
            return command;
        }

        private static void InsertAttachment(
            SQLiteConnection connection,
            SQLiteTransaction transaction,
            ModelChatAttachmentRecord attachment)
        {
            if (attachment == null || string.IsNullOrWhiteSpace(attachment.Id))
                return;

            const string sql =
                "INSERT INTO ModelAttachments (Id, RequestId, AttachmentType, FileName, FileExtension, ContentType, " +
                "SizeBytes, Width, Height) VALUES (@Id, @RequestId, @AttachmentType, @FileName, @FileExtension, " +
                "@ContentType, @SizeBytes, @Width, @Height);";
            using (SQLiteCommand command = new SQLiteCommand(sql, connection, transaction))
            {
                command.Parameters.AddWithValue("@Id", attachment.Id);
                command.Parameters.AddWithValue("@RequestId", attachment.RequestId ?? string.Empty);
                command.Parameters.AddWithValue("@AttachmentType", attachment.AttachmentType ?? string.Empty);
                command.Parameters.AddWithValue("@FileName", attachment.FileName ?? string.Empty);
                command.Parameters.AddWithValue("@FileExtension", attachment.FileExtension ?? string.Empty);
                command.Parameters.AddWithValue("@ContentType", attachment.ContentType ?? string.Empty);
                command.Parameters.AddWithValue("@SizeBytes", attachment.SizeBytes);
                command.Parameters.AddWithValue("@Width", DbValue(attachment.Width));
                command.Parameters.AddWithValue("@Height", DbValue(attachment.Height));
                command.ExecuteNonQuery();
            }
        }

        private static string GetCreateSessionsSql()
        {
            return "CREATE TABLE IF NOT EXISTS ModelChatSessions (" +
                "SessionId TEXT NOT NULL PRIMARY KEY, UserName TEXT NOT NULL, StartedAt TEXT NOT NULL, EndedAt TEXT, " +
                "RevitVersion INTEGER NOT NULL, RevitProcessId INTEGER NOT NULL, PluginVersion TEXT, ModelName TEXT, " +
                "ConnectionType TEXT);";
        }

        private static string GetCreateRequestsSql()
        {
            return "CREATE TABLE IF NOT EXISTS ModelChatRequests (" +
                "RequestId TEXT NOT NULL PRIMARY KEY, SessionId TEXT NOT NULL, RequestTime TEXT NOT NULL, ResponseTime TEXT, " +
                "UserQuestion TEXT NOT NULL, FinalAnswer TEXT, Status TEXT NOT NULL, Scenario TEXT NOT NULL, AiModelName TEXT, " +
                "RevitVersion INTEGER NOT NULL, RevitProcessId INTEGER NOT NULL, ModelName TEXT, ModelIdentity TEXT, " +
                "ViewName TEXT, DurationMs INTEGER, " +
                "UsedToolNamesJson TEXT NOT NULL DEFAULT '[]', CacheHitTokens INTEGER, CacheMissTokens INTEGER, " +
                "CompletionTokens INTEGER, Cost NUMERIC, ErrorMessage TEXT);";
        }

        private static string GetCreateAttachmentsSql()
        {
            return "CREATE TABLE IF NOT EXISTS ModelAttachments (" +
                "Id TEXT NOT NULL PRIMARY KEY, RequestId TEXT NOT NULL, AttachmentType TEXT NOT NULL, FileName TEXT NOT NULL, " +
                "FileExtension TEXT, ContentType TEXT, SizeBytes INTEGER NOT NULL, Width INTEGER, Height INTEGER);";
        }

        private static string GetCreateToolDefinitionsSql()
        {
            return "CREATE TABLE IF NOT EXISTS ToolDefinitions (" +
                "ToolId INTEGER NOT NULL PRIMARY KEY, ToolName TEXT NOT NULL UNIQUE, Description TEXT, OperationType TEXT NOT NULL, " +
                "IsReadOnly INTEGER NOT NULL, SupportsPagination INTEGER NOT NULL, PaginationUnit TEXT, DefaultPageSize INTEGER);";
        }

        private static string GetCreateFeedbackSql()
        {
            return "CREATE TABLE IF NOT EXISTS ModelAnswerFeedback (" +
                "FeedbackId TEXT NOT NULL PRIMARY KEY, RequestId TEXT NOT NULL UNIQUE, Rating TEXT NOT NULL, " +
                "ReasonCode TEXT, Comment TEXT);";
        }

        private static void ExecuteNonQuery(SQLiteConnection connection, SQLiteTransaction transaction, string sql)
        {
            using (SQLiteCommand command = new SQLiteCommand(sql, connection, transaction))
                command.ExecuteNonQuery();
        }

        private static object DbValue(object value)
        {
            return value ?? DBNull.Value;
        }

        private static object DbDateValue(DateTime? value)
        {
            return value.HasValue ? (object)ToDbDate(value.Value) : DBNull.Value;
        }

        private static string ToDbDate(DateTime value)
        {
            DateTime utc;
            if (value.Kind == DateTimeKind.Utc)
            {
                utc = value;
            }
            else if (value.Kind == DateTimeKind.Local)
            {
                utc = value.ToUniversalTime();
            }
            else
            {
                utc = TimeZoneInfo.ConvertTimeToUtc(value, TimeZoneInfo.Local);
            }

            return FormatMoscowDate(utc);
        }

        private static string FormatMoscowDate(DateTime utc)
        {
            DateTime normalizedUtc = utc.Kind == DateTimeKind.Utc
                ? utc
                : DateTime.SpecifyKind(utc, DateTimeKind.Utc);
            DateTime moscow = TimeZoneInfo.ConvertTimeFromUtc(normalizedUtc, MoscowTimeZone);
            return moscow.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static bool TryParseDbDateToUtc(string value, out DateTime utc)
        {
            if (DateTime.TryParseExact(
                value,
                "yyyy-MM-dd'T'HH:mm:ss.fff'Z'",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out utc))
            {
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);
                return true;
            }

            DateTime moscow;
            if (DateTime.TryParseExact(
                value,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out moscow))
            {
                moscow = DateTime.SpecifyKind(moscow, DateTimeKind.Unspecified);
                utc = TimeZoneInfo.ConvertTimeToUtc(moscow, MoscowTimeZone);
                return true;
            }

            utc = default(DateTime);
            return false;
        }

        private static TimeZoneInfo ResolveMoscowTimeZone()
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById("Russian Standard Time");
            }
            catch
            {
                return TimeZoneInfo.CreateCustomTimeZone(
                    "CoordinatorAI Moscow Time",
                    TimeSpan.FromHours(3),
                    "Moscow Time",
                    "Moscow Time");
            }
        }

        private static string SanitizeFileName(string value)
        {
            string result = string.IsNullOrWhiteSpace(value) ? "unknown_user" : value.Trim();
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
                result = result.Replace(invalidChar, '_');
            return result;
        }

        private static void EnsureFolder()
        {
            if (!Directory.Exists(DatabaseFolder))
                Directory.CreateDirectory(DatabaseFolder);
        }
    }

    internal static class DeepSeekUsageCostCalculator
    {
        public static bool TryCalculateUsd(
            string modelName,
            DateTime requestTimeUtc,
            long cacheHitTokens,
            long cacheMissTokens,
            long completionTokens,
            out decimal cost)
        {
            cost = 0m;
            string normalizedModel = (modelName ?? string.Empty).Trim().ToLowerInvariant();
            decimal cacheHitRate;
            decimal cacheMissRate;
            decimal outputRate;

            if (normalizedModel.Contains("v4-pro"))
            {
                cacheHitRate = 0.022m;
                cacheMissRate = 0.66m;
                outputRate = 1.98m;
            }
            else if (normalizedModel.Contains("v4-flash")
                || normalizedModel.Contains("deepseek-chat")
                || normalizedModel.Contains("deepseek-reasoner"))
            {
                cacheHitRate = 0.007m;
                cacheMissRate = 0.22m;
                outputRate = 0.66m;
            }
            else
            {
                return false;
            }

            if (IsPeakTime(requestTimeUtc))
            {
                cacheHitRate *= 2m;
                cacheMissRate *= 2m;
                outputRate *= 2m;
            }

            cost = ((cacheHitTokens * cacheHitRate)
                + (cacheMissTokens * cacheMissRate)
                + (completionTokens * outputRate)) / 1000000m;
            return true;
        }

        private static bool IsPeakTime(DateTime requestTimeUtc)
        {
            DateTime utc = requestTimeUtc.Kind == DateTimeKind.Utc
                ? requestTimeUtc
                : requestTimeUtc.ToUniversalTime();
            if (utc.DayOfWeek == DayOfWeek.Saturday || utc.DayOfWeek == DayOfWeek.Sunday)
                return false;

            TimeSpan time = utc.TimeOfDay;
            return (time >= TimeSpan.FromHours(1) && time < TimeSpan.FromHours(4))
                || (time >= TimeSpan.FromHours(6) && time < TimeSpan.FromHours(10));
        }
    }
}
