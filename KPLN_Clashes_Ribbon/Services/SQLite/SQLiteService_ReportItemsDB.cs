using Dapper;
using System.Data.SQLite;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    /// <summary>
    /// Личный сервис по работе с БД инстанса отчета
    /// </summary>
    public sealed partial class SQLiteService_ReportItemsDB : AbsSQLiteService
    {
        private const string _dbTableName = "Items";

        public SQLiteService_ReportItemsDB(string dbConnectionPath)
        {
            _dbConnectionPath = $@"Data Source={dbConnectionPath};Version=3;";
        }

        // Только данные: экземпляры WPF и Revit в фоновую запись не передаются.
        public sealed class ItemChange
        {
            public string Actor { get; set; }
            public string SessionId { get; set; }
            public int Id { get; set; }
            public int DepartmentId { get; set; }
            public string Comment { get; set; }
        }

        public sealed class SavedItem
        {
            public long Revision { get; set; }
            public int Id { get; set; }
            public int StatusId { get; set; }
            public int DelegatedDepartmentId { get; set; }
            public string Comments { get; set; }
        }

        public sealed class StatusCount
        {
            public int StatusId { get; set; }
            public long ItemCount { get; set; }
        }

        public sealed class ReportSummary
        {
            public long Revision { get; set; }
            public StatusCount[] Counts { get; set; }
            public KPItemStatus Status => GetSummaryStatus(Counts);
        }

        public static KPItemStatus GetSummaryStatus(IEnumerable<StatusCount> counts)
        {
            StatusCount[] rows = counts.Where(c => c.ItemCount > 0).ToArray();
            if (rows.Any(c => c.StatusId == (int)KPItemStatus.Opened))
                return KPItemStatus.Opened;
            if (rows.All(c => c.StatusId == (int)KPItemStatus.Closed))
                return KPItemStatus.Closed;
            if (rows.Any(c => c.StatusId == (int)KPItemStatus.Delegated))
                return KPItemStatus.Delegated;
            if (rows.All(c => c.StatusId == (int)KPItemStatus.Approved || c.StatusId == (int)KPItemStatus.Closed))
                return KPItemStatus.Approved;

            return KPItemStatus.Opened;
        }

        public ReportSummary ReadSummary()
        {
            EnsureChangeTracking();
            using (var connection = SQLiteRetryPolicy.CreateConnection(_dbConnectionPath))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted))
                {
                    var summary = new ReportSummary
                    {
                        Revision = connection.QuerySingle<long>("SELECT Revision FROM ReportRevision WHERE Id=1", transaction: transaction, commandTimeout: 1),
                        Counts = connection.Query<StatusCount>(
                            "SELECT StatusId, COUNT(*) AS ItemCount FROM Items WHERE ParentGroupId = -1 GROUP BY StatusId",
                            transaction: transaction, commandTimeout: 1).ToArray()
                    };
                    transaction.Commit();
                    return summary;
                }
            }
        }

        public Task<ReportSummary> ReadSummaryAsync() => SQLiteRetryPolicy.RunAsync(ReadSummary);

        public KPItemStatus ReadSummaryStatus() => ReadSummary().Status;

        public Task<KPItemStatus> ReadSummaryStatusAsync() => SQLiteRetryPolicy.RunAsync(ReadSummaryStatus);

        public Task<SavedItem[]> SaveChangesBatchAsync(ItemChange[] changes, KPItemStatus? status)
        {
            // Не позволяем вызывающему коду изменить пакет между попытками.
            ItemChange[] snapshot = changes.Select(c => new ItemChange
            {
                Id = c.Id, DepartmentId = c.DepartmentId, Comment = c.Comment,
                Actor = c.Actor, SessionId = c.SessionId
            }).ToArray();
            return SQLiteRetryPolicy.RunAsync(() => SaveChangesBatch(snapshot, status));
        }

        /// <summary>Сохраняет весь выбор либо откатывает его; ошибки передаются вызывающему коду.</summary>
        public SavedItem[] SaveChangesBatch(ItemChange[] changes, KPItemStatus? status)
        {
            if (changes.Length == 0)
                return new SavedItem[0];
            if (changes.Select(c => c.Id).Distinct().Count() != changes.Length)
                throw new ArgumentException("Пакет содержит повторяющиеся замечания.", nameof(changes));

            EnsureChangeTracking();

            // Подготовка запроса и параметров до захвата блокировки записи.
            string query = "UPDATE Items SET Comments = @Comment || " +
                "CASE WHEN COALESCE(Comments, '') = '' THEN '' ELSE @Separator || Comments END";
            if (status.HasValue)
                query += ", StatusId = @StatusId, DelegatedDepartmentId = @DepartmentId";
            query += " WHERE Id = @Id";
            var parameters = changes.Select(change => new
            {
                change.Id, change.Comment, change.DepartmentId,
                StatusId = (int?)status,
                Separator = ClashesMainCollection.StringSeparatorItem
            }).ToArray();

            using (var connection = SQLiteRetryPolicy.CreateConnection(_dbConnectionPath))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    connection.Execute("UPDATE ReportWriterContext SET Actor=@Actor, SessionId=@SessionId WHERE Id=1",
                        new { Actor = changes[0].Actor ?? string.Empty, SessionId = changes[0].SessionId ?? string.Empty },
                        transaction, commandTimeout: 1);
                    var saved = new List<SavedItem>();
                    foreach (var parameter in parameters)
                    {
                        // Добавляем к текущей истории в БД, а не к копии из открытого окна.
                        int affected = connection.Execute(query, parameter, transaction, commandTimeout: 1);
                        if (affected != 1)
                            throw new InvalidOperationException("Замечание отсутствует в БД: " + parameter.Id);
                    }

                    // Читаем только затронутые строки небольшими пакетами, без изображений.
                    for (int offset = 0; offset < changes.Length; offset += 500)
                        saved.AddRange(connection.Query<SavedItem>(
                            "SELECT i.Id, i.StatusId, i.DelegatedDepartmentId, i.Comments, " +
                            "COALESCE((SELECT MAX(c.Revision) FROM ReportChanges c WHERE c.ItemId=i.Id),0) AS Revision FROM Items i WHERE i.Id IN @Ids",
                            new { Ids = changes.Skip(offset).Take(500).Select(c => c.Id).ToArray() },
                            transaction, commandTimeout: 1));

                    SavedItem[] result = saved.ToArray();
                    connection.Execute("UPDATE ReportWriterContext SET Actor='', SessionId='' WHERE Id=1", transaction: transaction, commandTimeout: 1);
                    transaction.Commit();
                    return result;
                }
            }
        }

        #region Create
        /// <summary>
        /// Создание БД для хранения данных ReportItem
        /// </summary>
        public void CreateDbFile_ByReports() =>
            ExecuteNonQuery(
                    $"CREATE TABLE {_dbTableName} " +
                        $"({nameof(ReportItem.Id)} INTEGER PRIMARY KEY, " +
                        $"{nameof(ReportItem.ReportGroupId)} INTEGER, " +
                        $"{nameof(ReportItem.ReportId)} INTEGER, " +
                        $"{nameof(ReportItem.Name)} TEXT, " +
                        $"{nameof(ReportItem.Image)} BLOB, " +
                        $"{nameof(ReportItem.Element_1_DocName)} TEXT, " +
                        $"{nameof(ReportItem.Element_1_Id)} INTEGER NOT NULL DEFAULT -1, " +
                        $"{nameof(ReportItem.Element_1_Info)} TEXT, " +
                        $"{nameof(ReportItem.Element_2_DocName)} TEXT, " +
                        $"{nameof(ReportItem.Element_2_Id)} INTEGER NOT NULL DEFAULT -1, " +
                        $"{nameof(ReportItem.Element_2_Info)} TEXT, " +
                        $"{nameof(ReportItem.Point)} TEXT, " +
                        $"{nameof(ReportItem.StatusId)} INTEGER NOT NULL DEFAULT 1, " +
                        $"{nameof(ReportItem.ReportParentGroupComments)} TEXT, " +
                        $"{nameof(ReportItem.ReportItemComments)} TEXT, " +
                        $"{nameof(ReportItem.Comments)} TEXT, " +
                        $"{nameof(ReportItem.ParentGroupId)} INTEGER NOT NULL DEFAULT -1, " +
                        $"{nameof(ReportItem.DelegatedDepartmentId)} INTEGER)");

        /// <summary>
        /// Опубликовать данные ReportItem
        /// </summary>
        /// <param name="reports">Коллекция отчетов</param>
        public void PostNewItems_ByReports(ObservableCollection<ReportItem> reports) =>
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                    $"({nameof(ReportItem.Id)}, " +
                    $"{nameof(ReportItem.ReportGroupId)}, " +
                    $"{nameof(ReportItem.ReportId)}, " +
                    $"{nameof(ReportItem.Name)}, " +
                    $"{nameof(ReportItem.Image)}, " +
                    $"{nameof(ReportItem.Element_1_DocName)}, " +
                    $"{nameof(ReportItem.Element_1_Id)}, " +
                    $"{nameof(ReportItem.Element_1_Info)}, " +
                    $"{nameof(ReportItem.Element_2_DocName)}, " +
                    $"{nameof(ReportItem.Element_2_Id)}, " +
                    $"{nameof(ReportItem.Element_2_Info)}, " +
                    $"{nameof(ReportItem.Point)}, " +
                    $"{nameof(ReportItem.ReportParentGroupComments)}, " +
                    $"{nameof(ReportItem.ReportItemComments)}, " +
                    $"{nameof(ReportItem.StatusId)}, " +
                    $"{nameof(ReportItem.ParentGroupId)}) " +
                "VALUES " +
                    "(@Id, @ReportGroupId, @ReportId, @Name, @Image, @Element_1_DocName, @Element_1_Id, @Element_1_Info, " +
                    "@Element_2_DocName, @Element_2_Id, @Element_2_Info, @Point, @ReportParentGroupComments, " +
                    "@ReportItemComments, @StatusId, @ParentGroupId)",
                reports);
        #endregion

        #region Read
        /// <summary>
        /// Подготовить и Получить коллекцию всех ReportItem
        /// </summary>
        public ReportItem[] GetAllReporItems()
        {
            ReportItem[] repItems = ExecuteQuery<ReportItem>($"SELECT * FROM {_dbTableName}").ToArray();

            ReportItem[] groupItems = repItems.Where(i => i.ParentGroupId == -1).ToArray();
            ReportItem[] subItems = repItems.Where(i => i.ParentGroupId != -1).ToArray();

            // Врестаю структуру по одиночкам (нет ParentGroupId)
            foreach (ReportItem ri in groupItems)
            {
                // Добавляю сам себя, чтобы в окне отразились данные (если id по эл-там -1, значит это группа - такая структура для сохранения работы со старыми БД)
                if (ri.Element_2_Id != -1 && ri.Element_2_Id != -1)
                    ri.SubElements.Add(ri);

                // Добавляю остальные субэлементы
                ReportItem[] riSubItems = subItems.Where(si => si.ParentGroupId == ri.Id).ToArray();
                foreach (ReportItem si in riSubItems)
                {
                    ri.SubElements.Add(si);
                }
            }

            return groupItems;
        }

        /// <summary>
        /// Получить массив байт изображения ReportItem
        /// </summary>
        /// <param name="reportItem">ReportItem для поиска</param>
        /// <returns></returns>
        public byte[] GetImageBytes_ByItem(ReportItem reportItem) =>
            ExecuteQuery<byte[]>(
                $"SELECT {nameof(ReportItem.Image)} FROM {_dbTableName} " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}")
            .FirstOrDefault();

        /// <summary>
        /// Получить комментарии для ReportItem
        /// </summary>
        /// <param name="reportItem">ReportItem для поиска</param>
        /// <returns></returns>
        public string GetComment_ByReportItem(ReportItem reportItem) =>
            ExecuteQuery<string>(
                $"SELECT {nameof(ReportItem.Comments)} FROM {_dbTableName} " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}")
            .FirstOrDefault();
        #endregion

        #region Update
        /// <summary>
        /// Записать статус и отдел по Id-отчета (ReportItem)
        /// </summary>
        /// <param name="status">Статус для записи</param>
        /// <param name="departmentId">Id-отдела</param>
        /// <param name="reportItem">ReportItem для поиска</param>
        public void SetStatusAndDepartment_ByReportItem(ClashesMainCollection.KPItemStatus status, ReportItem reportItem) =>
            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.StatusId)}={(int)status}, {nameof(ReportItem.DelegatedDepartmentId)}={reportItem.DelegatedDepartmentId} " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}");

        /// <summary>
        /// Записать статус по Id-отчета (ReportItem)
        /// </summary>
        /// <param name="status">Статус для записи</param>
        /// <param name="reportItem">ReportItem для поиска</param>
        public void SetStatusId_ByReportItem(ClashesMainCollection.KPItemStatus status, ReportItem reportItem) =>
            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.StatusId)}={(int)status} " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}");

        /// <summary>
        /// Добавить комментарий для отчета (ReportItem)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="reportItem">ReportItem для поиска</param>
        public void SetComment_ByReportItem(string message, ReportItem reportItem)
        {
            List<string> value_parts = new List<string>
            {
                new ReportItemComment(message).ToString()
            };

            foreach (ReportItemComment comment in reportItem.CommentCollection)
            {
                value_parts.Add(comment.ToString());
            }
            string decorateMsg = string.Join(ClashesMainCollection.StringSeparatorItem, value_parts);

            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.Comments)}='{decorateMsg}'" +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}");
        }



        /// <summary>
        /// Добавляет комментарий в начало (верх) списка для ReportItem, перезаписывает дату на текущую. Формат encodedLine: user~dept~dd.MM.yyyy HH:mm:ss~dept2~text
        /// </summary>
        public void PrependEncodedComment_ByReportItem(string encodedLine, ReportItem reportItem, bool overrideDateToNow)
        {
            if (reportItem == null || string.IsNullOrWhiteSpace(encodedLine))
                return;

            string finalLine = encodedLine;
            if (overrideDateToNow)
            {
                var parts = encodedLine.Split('~');
                if (parts.Length >= 5)
                {
                    parts[2] = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                    finalLine = string.Join("~", parts);
                }
            }

            string EscapeSql(string s) => s?.Replace("'", "''") ?? string.Empty;

            var value_parts = new List<string>
            {
                EscapeSql(finalLine)
            };

            _ = reportItem.Comments;
            foreach (ReportItemComment comment in (reportItem.CommentCollection ?? new System.Collections.ObjectModel.ObservableCollection<ReportItemComment>()))
            {
                value_parts.Add(EscapeSql(comment.ToString()));
            }

            string decorateMsg = string.Join(ClashesMainCollection.StringSeparatorItem, value_parts);

            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.Comments)}='{decorateMsg}' " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}");
        }

        /// <summary>
        /// Обновляет делегированный отдел и статус айтема одной командой UPDATE.
        /// </summary>
        public void UpdateDelegationAndStatus_ByReportItem(ReportItem reportItem, int delegatedDeptId, int statusId)
        {
            if (reportItem == null) return;
            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.DelegatedDepartmentId)}={delegatedDeptId}, {nameof(ReportItem.StatusId)}={statusId} " +
                $"WHERE {nameof(ReportItem.Id)}={reportItem.Id}");
        }

        #endregion

        #region Delete
        #endregion
    }
}
