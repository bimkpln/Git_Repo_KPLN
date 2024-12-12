using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System;
using System.Collections.Generic;
using System.IO;
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

        #region Create
        // Создается вручную напрямую в БД
        #endregion

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

        /// <summary>
        /// Получить активный проект из БД по открытому проекту Ревит
        /// </summary>
        /// <param name="fileName">Имя открытого файла Ревит</param>
        public DBProject GetDBProject_ByRevitDocFileName(string fileName) => 
            GetDBProjects()
            .FirstOrDefault(p => 
                fileName.Contains(p.MainPath) 
                || fileName.Contains(p.RevitServerPath)
                || fileName.Contains(p.RevitServerPath2)
                || fileName.Contains(p.RevitServerPath3)
                || fileName.Contains(p.RevitServerPath4));
        #endregion
    }
}
