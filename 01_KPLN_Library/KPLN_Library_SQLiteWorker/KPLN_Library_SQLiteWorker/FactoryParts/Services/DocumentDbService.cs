using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Documents в БД
    /// </summary>
    public class DocumentDbService : DbService
    {
        internal DocumentDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        /// <summary>
        /// Создать новый документ в БД
        /// </summary>
        public void CreateDBDocument(DBDocument dBDocument)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                $"({nameof(DBDocument.CentralPath)}, {nameof(DBDocument.ProjectId)}, {nameof(DBDocument.SubDepartmentId)}, {nameof(DBDocument.LastChangedUserId)}, {nameof(DBDocument.LastChangedData)}) " +
                $"VALUES (@{nameof(DBDocument.CentralPath)}, @{nameof(DBDocument.ProjectId)}, @{nameof(DBDocument.SubDepartmentId)}, @{nameof(DBDocument.LastChangedUserId)}, @{nameof(DBDocument.LastChangedData)});",
            dBDocument);
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ документов
        /// </summary>
        public IEnumerable<DBDocument> GetDBDocuments() =>
            ExecuteQuery<DBDocument>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить документы по id проекта и отдела
        /// </summary>
        public IEnumerable<DBDocument> GetDBDocuments_ByPrjIdAndSubDepId(int projectId, int subDepartmentId) =>
            ExecuteQuery<DBDocument>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBDocument.ProjectId)}='{projectId}'" +
                $"AND {nameof(DBDocument.SubDepartmentId)}='{subDepartmentId}';");

        /// <summary>
        /// Получить документы по пути к файлу
        /// </summary>
        public DBDocument GetDBDocuments_ByFileFullPath(string fileFullPath) =>
            ExecuteQuery<DBDocument>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBDocument.CentralPath)}='{fileFullPath}';")
            .FirstOrDefault();
        #endregion

        #region Update
        /// <summary>
        /// Обновить статус IsClosed документа по статусу проекта
        /// </summary>
        public void UpdateDBDocument_IsClosedByProject(DBProject dbProject) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                  $"SET {nameof(DBDocument.IsClosed)}='{dbProject.IsClosed}' WHERE {nameof(DBDocument.ProjectId)}='{dbProject.Id}';");

        /// <summary>
        /// Обновить данные по последнему изменению документа (LastChangedUserId, LastChangedData)
        /// </summary>
        public void UpdateDBDocument_LastChangedData(DBDocument dBDocument, DBUser dBUser, string data) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                $"SET {nameof(DBDocument.LastChangedUserId)}='{dBUser.Id}', {nameof(DBDocument.LastChangedData)}='{data}' " +
                $"WHERE {nameof(DBDocument.Id)}='{dBDocument.Id}';");
        #endregion
    }
}
