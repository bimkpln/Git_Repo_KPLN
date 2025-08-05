using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace KPLN_Clashes_Ribbon.Services.SQLite
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
        #endregion

        #region Delete
        #endregion
    }
}
