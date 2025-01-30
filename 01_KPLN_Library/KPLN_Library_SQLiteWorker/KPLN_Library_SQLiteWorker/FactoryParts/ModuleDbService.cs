using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class ModuleDbService : DbService
    {
        internal ModuleDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        #endregion

        #region Read
        /// <summary>
        /// Получить модуль по имени папки, в которой расположена dll (она совпадает с именем модуля)
        /// </summary>
        public DBModule GetDBModule_ByFiDirName(string dirName) =>
            ExecuteQuery<DBModule>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBModule.Name)} = '{dirName}';")
            .FirstOrDefault();
        #endregion

        #region Update
        #endregion

        #region Delete
        #endregion
    }
}
