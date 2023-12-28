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
        internal DocumentDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }
 
        /// <summary>
        /// Создать новый документ в БД
        /// </summary>
        public void CreateDBDocument(DBDocument dBDocument)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbTableName} " +
                $"({nameof(DBDocument.Name)}, {nameof(DBDocument.FullPath)}, {nameof(DBDocument.ProjectId)}, {nameof(DBDocument.SubDepartmentId)}) " +
                $"VALUES (@{nameof(DBDocument.Name)}, @{nameof(DBDocument.FullPath)}, @{nameof(DBDocument.ProjectId)}, @{nameof(DBDocument.SubDepartmentId)});",
            dBDocument);
        }

        /// <summary>
        /// Получить документы по id проекта и отдела
        /// </summary>
        public IEnumerable<DBDocument> GetDBDocuments_ByPrjIdAndSubDepId(int projectId, int subDepartmentId) =>
            ExecuteQuery<DBDocument>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBDocument.ProjectId)}='{projectId}'" +
                $"AND {nameof(DBDocument.SubDepartmentId)}='{subDepartmentId}';");


        /// <summary>
        /// Обновить статус IsClosed документа по статусу проекта
        /// </summary>
        public void UpdateDBDocument_IsClosedByProject(DBProject dbProject) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                  $"SET {nameof(DBDocument.IsClosed)}='{dbProject.IsClosed}' WHERE {nameof(DBDocument.ProjectId)}='{dbProject.Id}';");

    }
}
