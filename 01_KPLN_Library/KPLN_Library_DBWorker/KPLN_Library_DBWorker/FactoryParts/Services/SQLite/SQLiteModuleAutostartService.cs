using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using Dapper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом RevitDocExchanges в БД
    /// </summary>
    public class SQLiteModuleAutostartService : SQLiteService
    {
        // Как в Users: местное время до минуты. В сущностях и расчётах используем UTC.
        private const string _dateTextFormat = "yyyy.MM.dd_HH:mm";

        private const string _keyFilter = "UserId = @UserId AND RevitVersion = @RevitVersion AND ProjectId = @ProjectId " +
            "AND ModuleId = @ModuleId AND DBTableName = @DBTableName AND DBTableKeyId = @DBTableKeyId";

        internal SQLiteModuleAutostartService(string connectionString, DBEnumerator dbEnumerator)
            : base(new SQLiteConnectionStringBuilder(connectionString) { DateTimeKind = DateTimeKind.Utc }.ConnectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Сохранить назначения с обязательным расписанием.
        /// </summary>
        public void BulkCreateDBModuleAutostarts(IEnumerable<DBModuleAutostart> dBModuleAutostarts) =>
            SaveDBModuleAutostarts(dBModuleAutostarts, Enumerable.Empty<DBModuleAutostart>());
        #endregion

        #region Read
        /// <summary>
        /// Получить конфигурации по нужному модулю для нужного пользователя 
        /// </summary>
        public IEnumerable<DBModuleAutostart> GetDBModuleAutostarts_ByUserAndRVersionAndTable(int userId, int rVersion, string tableName) =>
            ReadAssignments("UserId = @UserId AND RevitVersion = @RevitVersion AND DBTableName = @DBTableName",
                new { UserId = userId, RevitVersion = rVersion, DBTableName = tableName });

        /// <summary>
        /// Получить конфигурации по нужному модулю для нужного пользователя по нужному проекту для нужной таблицы конфигураций
        /// </summary>
        public IEnumerable<DBModuleAutostart> GetDBModuleAutostartsByUserAndRVersionAndPrjIdAndTable(int userId, int rVersion, int prjId, int moduleId, string tableName) =>
            ReadAssignments("UserId = @UserId AND RevitVersion = @RevitVersion AND ProjectId = @ProjectId " +
                "AND ModuleId = @ModuleId AND DBTableName = @DBTableName",
                new { UserId = userId, RevitVersion = rVersion, ProjectId = prjId, ModuleId = moduleId, DBTableName = tableName });

        private DBModuleAutostart[] ReadAssignments(string filter, object parameters)
        {
            using (var connection = CreateConnection(_connectionString))
            {
                connection.Open();
                return ReadAssignments(connection, null, filter, parameters);
            }
        }

        private DBModuleAutostart[] ReadAssignments(IDbConnection connection, IDbTransaction transaction, string filter, object parameters)
        {
            var columns = new HashSet<string>(connection.Query<string>(
                "SELECT name FROM pragma_table_info(@TableName);", new { TableName = _dbTableName }, transaction), StringComparer.OrdinalIgnoreCase);
            // CAST исключает неявный перевод дат драйвером/Dapper в местное время.
            // Одинаковое чтение для TEXT и уже созданных столбцов с объявлением DATETIME.
            string start = columns.Contains(nameof(DBModuleAutostart.StartDateUtc)) ? "CAST(StartDateUtc AS TEXT)" : "NULL";
            string last = columns.Contains(nameof(DBModuleAutostart.LastExportDateUtc)) ? "CAST(LastExportDateUtc AS TEXT)" : "NULL";
            // До разделения дат LastExportDateUtc означал старт, а не завершение.
            string lastStart = columns.Contains(nameof(DBModuleAutostart.LastStartDateUtc)) ? "CAST(LastStartDateUtc AS TEXT)" : last;
            if (!columns.Contains(nameof(DBModuleAutostart.LastStartDateUtc)))
                last = "NULL";
            string recipient = columns.Contains(nameof(DBModuleAutostart.NotificationUserId)) ? "NotificationUserId" : "NULL";
            string interval = columns.Contains(nameof(DBModuleAutostart.IntervalHours)) ? "IntervalHours" : "NULL";
            string skipWeekends = columns.Contains(nameof(DBModuleAutostart.SkipWeekends)) ? "COALESCE(SkipWeekends, 0)" : "0";
            AutostartRecord[] records = connection.Query<AutostartRecord>(
                "SELECT Id, UserId, RevitVersion, ProjectId, ModuleId, DBTableName, DBTableKeyId, " +
                $"{recipient} AS NotificationUserId, {interval} AS IntervalHours, {skipWeekends} AS SkipWeekends, {start} AS StartDateText, {lastStart} AS LastStartDateText, {last} AS LastExportDateText " +
                $"FROM {_dbTableName} WHERE {filter};", parameters, transaction).ToArray();

            foreach (AutostartRecord record in records)
            {
                record.StartDateUtc = ParseUtc(record.StartDateText, record.Id, nameof(DBModuleAutostart.StartDateUtc));
                record.LastStartDateUtc = ParseUtc(record.LastStartDateText, record.Id, nameof(DBModuleAutostart.LastStartDateUtc));
                record.LastExportDateUtc = ParseUtc(record.LastExportDateText, record.Id, nameof(DBModuleAutostart.LastExportDateUtc));
            }
            return records;
        }

        private sealed class AutostartRecord : DBModuleAutostart
        {
            public string StartDateText { get; set; }
            public string LastStartDateText { get; set; }
            public string LastExportDateText { get; set; }
        }

        private static DateTime? ParseUtc(string value, int assignmentId, string column)
        {
            if (value == null)
                return null;

            if (DateTime.TryParseExact(value, _dateTextFormat, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime local))
                return DateTime.SpecifyKind(local, DateTimeKind.Local).ToUniversalTime();

            // Ранее записанные ISO/SQLite-даты читаем с прежней семантикой UTC.
            // Единая точность до минуты не допускает повторного старта после округления истории.
            if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTimeOffset parsed))
                return ToMinute(parsed.UtcDateTime);

            throw new FormatException($"Автозапуск {assignmentId}: некорректная дата в {column}: '{value}'.");
        }

        private static DateTime AsUtc(DateTime value) => value.Kind == DateTimeKind.Local
            ? value.ToUniversalTime() : DateTime.SpecifyKind(value, DateTimeKind.Utc);

        private static DateTime ToMinute(DateTime value) => new DateTime(
            value.Ticks - value.Ticks % TimeSpan.TicksPerMinute, value.Kind);

        private static DynamicParameters GetWriteParameters(DBModuleAutostart item)
        {
            var parameters = new DynamicParameters(item);
            parameters.Add(nameof(DBModuleAutostart.SkipWeekends), item.SkipWeekends ? 1 : 0, DbType.Int32);
            // Явный DbType.String исключает преобразование формата драйвером SQLite.
            parameters.Add(nameof(DBModuleAutostart.StartDateUtc), FormatDateText(item.StartDateUtc), DbType.String);
            parameters.Add(nameof(DBModuleAutostart.LastStartDateUtc), FormatDateText(item.LastStartDateUtc), DbType.String);
            parameters.Add(nameof(DBModuleAutostart.LastExportDateUtc), FormatDateText(item.LastExportDateUtc), DbType.String);
            return parameters;
        }

        private static string FormatDateText(DateTime? value) => value.HasValue
            ? AsUtc(value.Value).ToLocalTime().ToString(_dateTextFormat, CultureInfo.InvariantCulture) : null;
        #endregion

        #region Update
        /// <summary>
        /// Сохранить выбранные назначения и удалить снятые в одной транзакции.
        /// Недостающие столбцы добавляются, но старые записи не получают расписание автоматически.
        /// </summary>
        public void SaveDBModuleAutostarts(IEnumerable<DBModuleAutostart> selected, IEnumerable<DBModuleAutostart> removed)
        {
            DBModuleAutostart[] selectedItems = selected.ToArray();
            DBModuleAutostart[] removedItems = removed.ToArray();
            if (selectedItems.Any(item => !item.HasSchedule || !item.NotificationUserId.HasValue))
                throw new ArgumentException("У каждого выбранного автозапуска должны быть дата и время старта, интервал от 12 часов с шагом 12 и получатель уведомления.");

            using (var connection = CreateConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction(System.Data.IsolationLevel.Serializable))
                {
                    EnsureScheduleColumns(connection, transaction);

                    foreach (DBModuleAutostart item in removedItems)
                        connection.Execute($"DELETE FROM {_dbTableName} WHERE {_keyFilter};", item, transaction);

                    foreach (DBModuleAutostart item in selectedItems)
                    {
                        DynamicParameters parameters = GetWriteParameters(item);
                        bool exists = connection.ExecuteScalar<int>(
                            $"SELECT COUNT(*) FROM {_dbTableName} WHERE {_keyFilter};", item, transaction) > 0;
                        if (exists)
                        {
                            // Открытый редактор не должен затирать историю запусков исполнителя.
                            connection.Execute($"UPDATE {_dbTableName} SET NotificationUserId = @NotificationUserId, " +
                                $"StartDateUtc = @StartDateUtc, IntervalHours = @IntervalHours, SkipWeekends = @SkipWeekends WHERE {_keyFilter};", parameters, transaction);
                        }
                        else
                        {
                            // Новое назначение (в том числе копия) не наследует историю выгрузок.
                            connection.Execute($"INSERT INTO {_dbTableName} " +
                                "(UserId, RevitVersion, ProjectId, ModuleId, DBTableName, DBTableKeyId, NotificationUserId, StartDateUtc, IntervalHours, SkipWeekends) " +
                                "VALUES (@UserId, @RevitVersion, @ProjectId, @ModuleId, @DBTableName, @DBTableKeyId, @NotificationUserId, @StartDateUtc, @IntervalHours, @SkipWeekends);",
                                parameters, transaction);
                        }
                    }

                    transaction.Commit();
                }
            }
        }
        #endregion

        /// <summary>Срез назначений на момент запуска Revit; записи без расписания пропускаются.</summary>
        public DBModuleAutostart[] GetDueDBModuleAutostarts(int userId, int rVersion, int moduleId, string tableName, DateTime launchTimeUtc) =>
            GetDBModuleAutostarts_ByUserAndRVersionAndTable(userId, rVersion, tableName)
                .Where(item => item.ModuleId == moduleId && item.IsDue(launchTimeUtc))
                .ToArray();

        /// <summary>
        /// Атомарно перепроверить срок и зарегистрировать начало попытки, включая неуспешную.
        /// NULL означает, что назначение уже забрал другой процесс, изменили или удалили.
        /// </summary>
        public DBModuleAutostart TryStartExport(DBModuleAutostart assignment, DateTime attemptTimeUtc)
        {
            attemptTimeUtc = AsUtc(attemptTimeUtc);
            using (var connection = CreateConnection(_connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction(IsolationLevel.Serializable))
                {
                    EnsureScheduleColumns(connection, transaction);
                    DBModuleAutostart current = ReadAssignments(connection, transaction,
                        $"Id = @Id AND {_keyFilter}", assignment).SingleOrDefault();
                    if (current == null || !current.IsDue(attemptTimeUtc))
                        return null;

                    current.LastStartDateUtc = ToMinute(attemptTimeUtc);
                    connection.Execute($"UPDATE {_dbTableName} SET LastStartDateUtc = @LastStartDateUtc WHERE Id = @Id;", GetWriteParameters(current), transaction);
                    transaction.Commit();
                    return current;
                }
            }
        }

        /// <summary>Завершить именно зарегистрированную попытку, не затирая историю более нового запуска.</summary>
        public bool CompleteExport(DBModuleAutostart assignment, DateTime completionTimeUtc)
        {
            if (assignment == null || !assignment.LastStartDateUtc.HasValue)
                throw new ArgumentException("Не зарегистрировано начало попытки выгрузки.");

            DateTime completed = ToMinute(AsUtc(completionTimeUtc));
            var parameters = GetWriteParameters(assignment);
            parameters.Add(nameof(DBModuleAutostart.LastExportDateUtc), FormatDateText(completed), DbType.String);

            using (var connection = CreateConnection(_connectionString))
            {
                connection.Open();
                int updated = connection.Execute($"UPDATE {_dbTableName} SET LastExportDateUtc = @LastExportDateUtc " +
                    $"WHERE Id = @Id AND {_keyFilter} AND LastStartDateUtc = @LastStartDateUtc;", parameters);
                if (updated == 0)
                    return false;

                assignment.LastExportDateUtc = completed;
                return true;
            }
        }

        private void EnsureScheduleColumns(IDbConnection connection, IDbTransaction transaction)
        {
            var columns = new HashSet<string>(connection.Query<string>(
                "SELECT name FROM pragma_table_info(@TableName);", new { TableName = _dbTableName }, transaction), StringComparer.OrdinalIgnoreCase);
            var required = new Dictionary<string, string>
            {
                { nameof(DBModuleAutostart.NotificationUserId), "INTEGER" },
                { nameof(DBModuleAutostart.StartDateUtc), "TEXT" },
                { nameof(DBModuleAutostart.IntervalHours), "INTEGER" },
                { nameof(DBModuleAutostart.LastExportDateUtc), "TEXT" },
                { nameof(DBModuleAutostart.LastStartDateUtc), "TEXT" },
            };
            foreach (var column in required)
            {
                if (!columns.Contains(column.Key))
                    connection.Execute($"ALTER TABLE {_dbTableName} ADD COLUMN {column.Key} {column.Value} NULL;", transaction: transaction);
            }

            if (!columns.Contains(nameof(DBModuleAutostart.SkipWeekends)))
                connection.Execute($"ALTER TABLE {_dbTableName} ADD COLUMN SkipWeekends INTEGER NOT NULL DEFAULT 0;", transaction: transaction);

            // Однократный перенос старой отметки начала. Время завершения старых попыток неизвестно.
            if (!columns.Contains(nameof(DBModuleAutostart.LastStartDateUtc)))
                connection.Execute($"UPDATE {_dbTableName} SET LastStartDateUtc = LastExportDateUtc, LastExportDateUtc = NULL;", transaction: transaction);
        }

        #region Delete
        /// <summary>
        /// Удаление сущности в БД
        /// </summary>
        public void DeleteDBModuleAutostarts(DBModuleAutostart dBModuleAutostart) =>
            ExecuteNonQuery($"DELETE FROM {_dbTableName} WHERE {_keyFilter};", dBModuleAutostart);
        #endregion
    }
}
