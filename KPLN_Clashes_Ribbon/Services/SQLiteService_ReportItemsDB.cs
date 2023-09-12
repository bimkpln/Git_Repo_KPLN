using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Services
{
    /// <summary>
    /// Личный сервис по работе с БД инстанса отчета
    /// </summary>
    public sealed class SQLiteService_ReportItemsDB : AbsSQLiteService
    {
        private const string _dbTableName = "Items";

        public SQLiteService_ReportItemsDB(string dbConnectionPath)
        {
            _dbConnectionPath = $@"Data Source={dbConnectionPath};Version=3;";
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
                        $"{nameof(ReportItem.Name)} TEXT, " +
                        $"{nameof(ReportItem.Image)} BLOB, " +
                        $"{nameof(ReportItem.Element_1_Info)} TEXT, " +
                        $"{nameof(ReportItem.Element_2_Info)} TEXT, " +
                        $"{nameof(ReportItem.Point)} TEXT, " +
                        $"{nameof(ReportItem.Status)} TEXT NOT NULL DEFAULT 'Opened', " +
                        $"{nameof(ReportItem.Comments)} TEXT, " +
                        $"{nameof(ReportItem.GroupId)} INTEGER, " +
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
                    $"{nameof(ReportItem.Name)}, " +
                    $"{nameof(ReportItem.Image)}, " +
                    $"{nameof(ReportItem.Element_1_Info)}, " +
                    $"{nameof(ReportItem.Element_2_Info)}, " +
                    $"{nameof(ReportItem.Point)}, " +
                    $"{nameof(ReportItem.Status)}) " +
                "VALUES " +
                    "(@Id, @ReportGroupId, @Name, @Image, " +
                    "@Element_1_Info, @Element_2_Info, @Point, @Status)",
                reports);
        #endregion

        #region Read
        #endregion

        #region Update
        /// <summary>
        /// Записать статус и отдел по Id-отчета (ReportInstance)
        /// </summary>
        /// <param name="status">Статус для записи</param>
        /// <param name="departmentId">Id-отдела</param>
        /// <param name="itemId">Id элемента (ReportInstance)</param>
        public void SetStatusAndDepartment_ByReportItemId(ClashesMainCollection.KPItemStatus status, int departmentId, int itemId) =>
            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.Status)}={status}, {nameof(ReportItem.DelegatedDepartmentId)}={departmentId} " +
                $"WHERE Id={itemId}");

        /// <summary>
        /// Записать статус по Id-отчета (ReportInstance)
        /// </summary>
        /// <param name="status">Статус для записи</param>
        /// <param name="itemId">Id элемента (ReportInstance)</param>
        public void SetStatus_ByReportItemId(ClashesMainCollection.KPItemStatus status, int itemId) =>
            ExecuteNonQuery(
                $"UPDATE {_dbTableName} " +
                $"SET {nameof(ReportItem.Status)}={status}" +
                $"WHERE Id={itemId}");
        #endregion

        #region Delete
        #endregion


    }
}
