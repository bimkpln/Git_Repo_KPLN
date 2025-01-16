using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System;
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
        public DBProject GetDBProject_ByRevitDocFileName(string fileName)
        {
            IEnumerable<DBProject> filteredPrjs = GetDBProjects()
                .Where(p =>
                    fileName.Contains(p.MainPath)
                    || fileName.Contains(p.RevitServerPath)
                    || fileName.Contains(p.RevitServerPath2)
                    || fileName.Contains(p.RevitServerPath3)
                    || fileName.Contains(p.RevitServerPath4));
            
            // Осуществяляю поиск по наиболее подходящему варианту
            return filteredPrjs
                .OrderByDescending(prj => GetMatchingSegmentsCount(fileName, prj.MainPath))
                .FirstOrDefault();
        }
        #endregion

        private static int GetMatchingSegmentsCount(string path1, string path2)
        {
            var segments1 = path1.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var segments2 = path2.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            int matchingCount = 0;

            for (int i = 0; i < Math.Min(segments1.Length, segments2.Length); i++)
            {
                if (segments1[i] == segments2[i])
                {
                    matchingCount++;
                }
                else
                {
                    break;
                }
            }

            return matchingCount;
        }
    }
}
