using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Projects в БД
    /// </summary>
    public class ProjectDbService : DbService
    {
        internal ProjectDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ проектов
        /// </summary>
        public IEnumerable<DBProject> GetDBProjects() =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить коллекцию ВСЕХ открытых проектов
        /// </summary>
        public IEnumerable<DBProject> GetDBProjects_Opened() =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProject.IsClosed)}='False';");

        /// <summary>
        /// Получить проект по Id
        /// </summary>
        public DBProject GetDBProject_ByProjectId(int id) => 
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProject.Id)}='{id}';")
            .FirstOrDefault();
        #endregion
    }
}
