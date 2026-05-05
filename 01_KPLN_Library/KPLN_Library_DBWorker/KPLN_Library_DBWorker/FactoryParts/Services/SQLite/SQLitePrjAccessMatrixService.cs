using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Collections.Generic;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class SQLitePrjAccessMatrixService : SQLiteService
    {
        internal SQLitePrjAccessMatrixService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы
        /// </summary>
        public IEnumerable<DBProjectsAccessMatrix> GetDBProjectMatrix() =>
            ExecuteQuery<DBProjectsAccessMatrix>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы по выбранному проекту
        /// </summary>
        public IEnumerable<DBProjectsAccessMatrix> GetDBProjectMatrix_ByProject(DBProject dBProject) =>
            ExecuteQuery<DBProjectsAccessMatrix>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProjectsAccessMatrix.ProjectId)}='{dBProject.Id}';");
        #endregion

        #region Update

        #endregion

        #region Delete
        #endregion
    }
}
