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
        internal ProjectDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }

        /// <summary>
        /// Получить коллекцию ВСЕХ проектов
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public IEnumerable<DBProject> GetDBProjects() => 
            ExecuteQuery<DBProject>($"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить проект по Id
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public DBProject GetDBProject_ByProjectId(int id) => 
            ExecuteQuery<DBProject>($"SELECT * FROM {_dbTableName} WHERE {nameof(DBProject.Id)}='{id}';").FirstOrDefault();
    }
}
