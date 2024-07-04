using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class ProjectMatrixDbService : DbService
    {
        internal ProjectMatrixDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы
        /// </summary>
        public IEnumerable<DBProjectMatrix> GetDBProjectMatrix() =>
            ExecuteQuery<DBProjectMatrix>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить коллекцию ВСЕХ единиц матрицы по выбранному проекту
        /// </summary>
        public IEnumerable<DBProjectMatrix> GetDBProjectMatrix_ByProject(DBProject dBProject) =>
            ExecuteQuery<DBProjectMatrix>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProjectMatrix.ProjectId)}='{dBProject.Id}';");
        #endregion

        #region Update

        #endregion

        #region Delete
        #endregion
    }
}
