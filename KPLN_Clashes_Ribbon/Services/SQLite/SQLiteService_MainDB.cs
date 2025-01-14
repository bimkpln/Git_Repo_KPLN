using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    /// <summary>
    /// Личный сервис по работе с общей БД (которая хранит ссылки на отчеты и основную информацию)
    /// </summary>
    public sealed class SQLiteService_MainDB : AbsSQLiteService
    {
        private static readonly string _userSystemName = CurrentDBUser.SystemName;

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
        public void PostReportGroups_NewGroupByProjectAndName(DBProject project, string name) =>
            ExecuteNonQuery(
                $"INSERT INTO {MainDB_Enumerator.ReportGroups} " +
                    $"({nameof(ReportGroup.ProjectId)}, " +
                    $"{nameof(ReportGroup.Name)}, " +
                    $"{nameof(ReportGroup.Status)}, " +
                    $"{nameof(ReportGroup.DateCreated)}, " +
                    $"{nameof(ReportGroup.UserCreated)}, " +
                    $"{nameof(ReportGroup.DateLast)}, " +
                    $"{nameof(ReportGroup.UserLast)}) " +
                "VALUES " +
                    $"({project.Id}, " +
                    $"'{name}', " +
                    $"'{KPItemStatus.New}', " +
                    $"'{CurrentTime}', " +
                    $"'{_userSystemName}', " +
                    $"'{CurrentTime}', " +
                    $"'{_userSystemName}')");

        /// <summary>
        /// Опубликовать новый отчет (Report) с привязкой к группе (ReportGroup)
        /// </summary>
        /// <param name="name">Имя группы отчета</param>
        /// <param name="group">Экземпляр класса-группы (ReportGroup)</param>
        /// <param name="groupDbFi">Файл БД (ReportInstance)</param>
        public void PostReport_NewReport_ByNameAndReportGroup(string name, ReportGroup group, FileInfo groupDbFi) =>
            ExecuteNonQuery(
                $"INSERT INTO {MainDB_Enumerator.Reports} " +
                    $"({nameof(Report.ReportGroupId)}, " +
                    $"{nameof(Report.Name)}, " +
                    $"{nameof(Report.Status)}, " +
                    $"{nameof(Report.PathToReportInstance)}, " +
                    $"{nameof(Report.DateCreated)}, " +
                    $"{nameof(Report.UserCreated)}, " +
                    $"{nameof(Report.DateLast)}, " +
                    $"{nameof(Report.UserLast)})" +
                "VALUES " +
                    $"({group.Id}, " +
                    $"'{name}', " +
                    $"'{KPItemStatus.New}', " +
                    $"'{groupDbFi.FullName}', " +
                    $"'{CurrentTime}', " +
                    $"'{_userSystemName}', " +
                    $"'{CurrentTime}', " +
                    $"'{_userSystemName}')");
        #endregion

        #region Read
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
        /// Получить коллекцию открытых Report для текущей группы отчетов
        /// </summary>
        /// <param name="project">Проект для поиска</param>
        public ObservableCollection<Report> GetReports_OpenedByProject(DBProject project) =>
            new ObservableCollection<Report>(
                ExecuteQuery<Report>(
                    $"SELECT * FROM {MainDB_Enumerator.Reports} " +
                    $"WHERE {nameof(Report.ReportGroupId)} " +
                    $"IN (" +
                        $"SELECT {nameof(Report.ReportGroupId)} FROM {MainDB_Enumerator.ReportGroups} " +
                        $"WHERE {nameof(ReportGroup.ProjectId)} = {project.Id}) " +
                    $"AND {nameof(Report.Status)} = '{Core.ClashesMainCollection.KPItemStatus.Opened}'"));

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
                $"SET {nameof(ReportGroup.DateLast)}='{CurrentTime}', {nameof(ReportGroup.UserLast)}='{_userSystemName}' " +
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
                    $"{nameof(Report.UserLast)}='{_userSystemName}', " +
                    $"{nameof(Report.Status)}='{mainStatus}' " +
                $"WHERE Id={reportId}");
        #endregion

        #region Delete
        /// <summary>
        /// Удалить отчет (Report) и БД с единицами отчета (ReportItem) по Report
        /// </summary>
        /// <param name="id">Id отчета для удаления</param>
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
