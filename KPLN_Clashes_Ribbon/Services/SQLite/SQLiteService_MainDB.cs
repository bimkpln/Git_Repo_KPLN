using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
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
                    $"'{DBMainService.CurrentDBUser.Name} {DBMainService.CurrentDBUser.Surname}', " +
                    $"'{CurrentTime}', " +
                    $"'{DBMainService.CurrentDBUser.Name} {DBMainService.CurrentDBUser.Surname}', " +
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
                UserCreated = DBMainService.CurrentDBUser.SystemName,
                DateLast = CurrentTime,
                UserLast = DBMainService.CurrentDBUser.SystemName
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
                $"SET {nameof(ReportGroup.DateLast)}='{CurrentTime}', {nameof(ReportGroup.UserLast)}='{DBMainService.CurrentDBUser.Name} {DBMainService.CurrentDBUser.Surname}' " +
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
                    $"{nameof(Report.UserLast)}='{DBMainService.CurrentDBUser.SystemName}', " +
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
