using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class SQLitePrjOSClashMatrixService : SQLiteService
    {
        internal SQLitePrjOSClashMatrixService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы
        /// </summary>
        public IEnumerable<DBProjectsIOSClashMatrix> GetDBProjectMatrix() =>
            ExecuteQuery<DBProjectsIOSClashMatrix>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы по выбранному проекту
        /// </summary>
        public IEnumerable<DBProjectsIOSClashMatrix> GetDBProjectMatrix_ByProject(DBProject dBProject) =>
            ExecuteQuery<DBProjectsIOSClashMatrix>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProjectsIOSClashMatrix.ProjectId)}='{dBProject.Id}';");
        #endregion

        #region Update

        #endregion

        #region Delete
        #endregion
    }
}
