using Dapper;
using System.Data.SQLite;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    /// <summary>
    /// Личный сервис по работе с общей БД (которая хранит ссылки на отчеты и основную информацию)
    /// </summary>
    public sealed class SQLiteService_MainDB : AbsSQLiteService
    {
        public SQLiteService_MainDB()
        {
            _dbConnectionPath = @"Data Source=Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Clashes_MainDB.db;Version=3;";
        }

        /// <summary>Метки обновляются один раз на пакет; комментарий не меняет статус отчета.</summary>
        public Task<bool> UpdateBatchMetadataAsync(int reportId, int groupId, long changeRevision,
            SQLiteService_ReportItemsDB.ReportSummary summary, string userName, string fullName)
        {
            return SQLiteRetryPolicy.RunAsync(() =>
            {
                UpdateBatchMetadata(reportId, groupId, changeRevision, summary, userName, fullName);
                return true;
            });
        }

        public void UpdateBatchMetadata(int reportId, int groupId, long changeRevision,
            SQLiteService_ReportItemsDB.ReportSummary summary, string userName, string fullName)
        {
            var parameters = new { ReportId = reportId, GroupId = groupId, Date = CurrentTime,
                UserName = userName, FullName = fullName, Status = summary?.Status.ToString(),
                ChangeRevision = changeRevision, StatusRevision = summary?.Revision };

            using (var connection = SQLiteRetryPolicy.CreateConnection(_dbConnectionPath))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    // Раздельные версии: новый комментарий не блокирует отложенную сводку статусов.
                    connection.Execute(@"CREATE TABLE IF NOT EXISTS ReportSyncState (
                        ReportId INTEGER PRIMARY KEY, StatusRevision INTEGER NOT NULL DEFAULT -1,
                        MetadataRevision INTEGER NOT NULL DEFAULT -1);
                        INSERT OR IGNORE INTO ReportSyncState(ReportId) VALUES(@ReportId);",
                        parameters, transaction, commandTimeout: 1);
                    if (connection.QuerySingle<int>("SELECT COUNT(*) FROM Reports WHERE Id=@ReportId", parameters, transaction, commandTimeout: 1) != 1)
                        throw new System.InvalidOperationException("Отчет или группа отсутствует в общей БД.");

                    if (summary != null)
                    {
                        int updated = connection.Execute(@"UPDATE Reports SET Status=@Status WHERE Id=@ReportId
                            AND @StatusRevision >= (SELECT StatusRevision FROM ReportSyncState WHERE ReportId=@ReportId)",
                            parameters, transaction, commandTimeout: 1);
                        if (updated == 1)
                            connection.Execute("UPDATE ReportSyncState SET StatusRevision=@StatusRevision WHERE ReportId=@ReportId",
                                parameters, transaction, commandTimeout: 1);
                    }

                    int metadataUpdated = connection.Execute(@"UPDATE Reports SET DateLast=@Date, UserLast=@UserName WHERE Id=@ReportId
                        AND @ChangeRevision >= (SELECT MetadataRevision FROM ReportSyncState WHERE ReportId=@ReportId)",
                        parameters, transaction, commandTimeout: 1);
                    if (metadataUpdated == 1)
                    {
                        if (connection.Execute("UPDATE ReportGroups SET DateLast=@Date, UserLast=@FullName WHERE Id=@GroupId",
                            parameters, transaction, commandTimeout: 1) != 1)
                            throw new System.InvalidOperationException("Группа отсутствует в общей БД.");
                        connection.Execute("UPDATE ReportSyncState SET MetadataRevision=@ChangeRevision WHERE ReportId=@ReportId",
                            parameters, transaction, commandTimeout: 1);
                    }

                    transaction.Commit();
                }
            }
        }

        #region Create
        /// <summary>
        /// Опубликовать новую группу для проекта
        /// </summary>
        /// <param name="name">Имя группы</param>
        /// <param name="project">Проект</param>
        public void PostReportGroups_NewGroupByProjectAndName(DBProject project, ReportGroup rGroup) =>
            ExecuteNonQuery(
                $"INSERT INTO {MainDB_Enumerator.ReportGroups} " +
                    $"({nameof(ReportGroup.ProjectId)}, " +
                    $"{nameof(ReportGroup.Name)}, " +
                    $"{nameof(ReportGroup.Status)}, " +
                    $"{nameof(ReportGroup.DateCreated)}, " +
                    $"{nameof(ReportGroup.UserCreated)}, " +
                    $"{nameof(ReportGroup.DateLast)}, " +
                    $"{nameof(ReportGroup.UserLast)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdAR)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdKR)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdOV)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdITP)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdVK)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdAUPT)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdEOM)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdSS)}, " +
                    $"{nameof(ReportGroup.BitrixTaskIdAV)})" +
                "VALUES " +
                    $"({project.Id}, " +
                    $"'{rGroup.Name}', " +
                    $"'{KPItemStatus.New}', " +
                    $"'{CurrentTime}', " +
                    $"'{SQLiteMainService.CurrentDBUser.Name} {SQLiteMainService.CurrentDBUser.Surname}', " +
                    $"'{CurrentTime}', " +
                    $"'{SQLiteMainService.CurrentDBUser.Name} {SQLiteMainService.CurrentDBUser.Surname}', " +
                    $"'{rGroup.BitrixTaskIdAR}', " +
                    $"'{rGroup.BitrixTaskIdKR}', " +
                    $"'{rGroup.BitrixTaskIdOV}', " +
                    $"'{rGroup.BitrixTaskIdITP}', " +
                    $"'{rGroup.BitrixTaskIdVK}', " +
                    $"'{rGroup.BitrixTaskIdAUPT}', " +
                    $"'{rGroup.BitrixTaskIdEOM}', " +
                    $"'{rGroup.BitrixTaskIdSS}', " +
                    $"'{rGroup.BitrixTaskIdAV}')");

        /// <summary>
        /// Опубликовать новый отчет (Report) с привязкой к группе (ReportGroup)
        /// </summary>
        /// <param name="name">Имя группы отчета</param>
        /// <param name="group">Экземпляр класса-группы (ReportGroup)</param>
        /// <param name="groupDbFi">Файл БД (ReportInstance)</param>
        /// <param name="bitrTaskId_1">ID задачи из битрикс №1</param>
        /// <param name="bitrTaskId_2">ID задачи из битрикс №2</param>
        /// <returns>ID созданного Report</returns>
        public int PostReport_NewReport_ByNameAndReportGroup(string name, ReportGroup group, FileInfo groupDbFi)
        {
            string query =
                $"INSERT INTO {MainDB_Enumerator.Reports} " +
                $"({nameof(Report.ReportGroupId)}, " +
                $"{nameof(Report.Name)}, " +
                $"{nameof(Report.Status)}, " +
                $"{nameof(Report.PathToReportInstance)}, " +
                $"{nameof(Report.DateCreated)}, " +
                $"{nameof(Report.UserCreated)}, " +
                $"{nameof(Report.DateLast)}, " +
                $"{nameof(Report.UserLast)}) " +
                "VALUES " +
                $"(@GroupId, @Name, @Status, @Path, @DateCreated, @UserCreated, @DateLast, @UserLast); " +
                "SELECT last_insert_rowid();";

            var parameters = new
            {
                GroupId = group.Id,
                Name = name,
                Status = KPItemStatus.New.ToString(),
                Path = groupDbFi.FullName,
                DateCreated = CurrentTime,
                UserCreated = SQLiteMainService.CurrentDBUser.SystemName,
                DateLast = CurrentTime,
                UserLast = SQLiteMainService.CurrentDBUser.SystemName
            };

            return ExecuteInsertWithId(query, parameters);
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ReportGroup по ID
        /// </summary>
        public ReportGroup GetReportGroup_ById(int id) =>
            ExecuteQuery<ReportGroup>(
                $"SELECT * FROM {MainDB_Enumerator.ReportGroups} " +
                $"WHERE {nameof(ReportGroup.Id)} = {id}")
            .FirstOrDefault();

        /// <summary>
        /// Получить коллекцию ReportGroup для текущего проекта
        /// </summary>
        /// <param name="project">Проект для поиска</param>
        public ObservableCollection<ReportGroup> GetReportGroups_ByDBProject(DBProject project) =>
            new ObservableCollection<ReportGroup>(
                ExecuteQuery<ReportGroup>(
                    $"SELECT * FROM {MainDB_Enumerator.ReportGroups} " +
                    $"WHERE {nameof(ReportGroup.ProjectId)} = {project.Id}"));

        /// <summary>
        /// Получить коллекцию ТОЛЬКО открытых ReportGroup для текущего проекта
        /// </summary>
        /// <param name="project">Проект для поиска</param>
        public ObservableCollection<ReportGroup> GetReportGroups_ByDBProjectANDNotClosed(DBProject project) =>
            new ObservableCollection<ReportGroup>(
                ExecuteQuery<ReportGroup>(
                    $"SELECT * FROM {MainDB_Enumerator.ReportGroups} " +
                    $"WHERE {nameof(ReportGroup.ProjectId)} = {project.Id} " +
                    $"AND {nameof(ReportGroup.Status)} != '{KPItemStatus.Closed}'"));

        /// <summary>
        /// Получить коллекция Report для текущей группы отчетов (ReportGroup)
        /// </summary>
        /// <param name="reportGroupId">ID группы отчетов</param>
        public ObservableCollection<Report> GetReports_ByReportGroupId(int reportGroupId) =>
            new ObservableCollection<Report>(
                ExecuteQuery<Report>(
                    $"SELECT * FROM {MainDB_Enumerator.Reports} " +
                    $"WHERE {nameof(Report.ReportGroupId)} = {reportGroupId}"));
        #endregion

        #region Update
        /// <summary>
        /// Обновить статус элемента для текущей таблицы
        /// </summary>
        /// <param name="status">Статус для записи</param>
        /// <param name="dB_Enumerator">Таблица</param>
        /// <param name="itemId">Id элемента (ключ)</param>
        public void UpdateItemStatus_ByTableAndItemId(KPItemStatus status, MainDB_Enumerator dB_Enumerator, int itemId) =>
            ExecuteNonQuery(
                $"UPDATE {dB_Enumerator} " +
                $"SET Status='{status}' " +
                $"WHERE Id={itemId}");

        /// <summary>
        /// Обновить в группу отчетов меток последнего изменения по Id-группы
        /// </summary>
        /// <param name="groupId">Id-группы</param>
        public void UpdateReportGroup_MarksLastChange_ByGroupId(int groupId) =>
            ExecuteNonQuery(
                $"UPDATE {MainDB_Enumerator.ReportGroups} " +
                $"SET {nameof(ReportGroup.DateLast)}='{CurrentTime}', {nameof(ReportGroup.UserLast)}='{SQLiteMainService.CurrentDBUser.Name} {SQLiteMainService.CurrentDBUser.Surname}' " +
                $"WHERE Id={groupId}");

        /// <summary>
        /// Обновить в отчет меток последнего изменения по Id-группы и основному статусу ReportInstance в отчете
        /// </summary>
        /// <param name="reportId">Id очтета</param>
        /// <param name="mainStatus">Ключевой статус инстансов отчета (ReportInstance)</param>
        public void UpdateReport_MarksLastChange_ByIdAndMainRepInstStatus(int reportId, KPItemStatus mainStatus) =>
            ExecuteNonQuery(
                $"UPDATE {MainDB_Enumerator.Reports} " +
                $"SET {nameof(Report.DateLast)}='{CurrentTime}', " +
                    $"{nameof(Report.UserLast)}='{SQLiteMainService.CurrentDBUser.SystemName}', " +
                    $"{nameof(Report.Status)}='{mainStatus}' " +
                $"WHERE Id={reportId}");
        #endregion

        #region Delete
        /// <summary>
        /// Удалить отчет (Report) и БД с единицами отчета (ReportItem) по Report
        /// </summary>
        /// <param name="report">Отчет для удаления</param>
        public void DeleteReportAndReportItems_ByReportId(Report report)
        {
            // Удаляю строку из БД (Report).
            ExecuteNonQuery(
                $"DELETE FROM {MainDB_Enumerator.Reports} " +
                $"WHERE {nameof(Report.Id)} = {report.Id}");
            
            // Удаляю БД с единицами отчетов (ReportItem)
            Task repInstDeleteWorker = Task.Run(() =>
            {
                string repInstDbPath = report.PathToReportInstance;
                File.Delete(repInstDbPath);
            });
        }

        /// <summary>
        /// Удалить группу отчетов (ReportGroup), отчет (Report) и БД с единицами отчета (ReportItem) по id группы. 
        /// Удаление отчетов (Report) - лежит на БД (ON DELETE CASCADE)
        /// </summary>
        /// <param name="id">Id-группы для удаления</param>
        public void DeleteReportGroupAndReportsAndReportItems_ByReportGroupId(int id)
        {
            // Удаляю БД с единицами отчетов (ReportItem)
            ObservableCollection<Report> reports = GetReports_ByReportGroupId(id);
            Task repInstDeleteWorker = Task.Run(() =>
            {
                foreach (Report report in reports)
                {
                    string repInstDbPath = report.PathToReportInstance;
                    if (File.Exists(repInstDbPath))
                        File.Delete(repInstDbPath);
                }
            });

            // Удаляю строку из БД (ReportGroups). ON DELETE CASCADE - в БД не работает в запросах из кода!!!
            ExecuteNonQuery(
                $"DELETE FROM {MainDB_Enumerator.ReportGroups} " +
                $"WHERE {nameof(ReportGroup.Id)} = {id}");
            
            // Удаляю строки из БД (Report).
            ExecuteNonQuery(
                $"DELETE FROM {MainDB_Enumerator.Reports} " +
                $"WHERE {nameof(Report.ReportGroupId)} = {id}");
        }
        #endregion
    }
}
