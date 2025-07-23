using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

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
        /// <summary>
        /// Создать новый проект
        /// </summary>
        public Task<int> CreateDBDocument(DBProject dBProject)
        {
            // Дефолтное значение для пустых полей (для РС - это вполне нормально)
            if (string.IsNullOrWhiteSpace(dBProject.RevitServerPath))
                dBProject.RevitServerPath = "Empty";
            if (string.IsNullOrWhiteSpace(dBProject.RevitServerPath2))
                dBProject.RevitServerPath2 = "Empty";
            if (string.IsNullOrWhiteSpace(dBProject.RevitServerPath3))
                dBProject.RevitServerPath3 = "Empty";
            if (string.IsNullOrWhiteSpace(dBProject.RevitServerPath4))
                dBProject.RevitServerPath4 = "Empty";

            return Task<int>.Run(() =>
            {
                int createdPrjId = ExecuteQuery<int>(
                    $"INSERT INTO {_dbTableName} " +
                    $"({nameof(DBProject.Name)}, {nameof(DBProject.Code)}, {nameof(DBProject.Stage)}, {nameof(DBProject.RevitVersion)}, {nameof(DBProject.MainPath)}, " +
                        $"{nameof(DBProject.RevitServerPath)}, {nameof(DBProject.RevitServerPath2)}, {nameof(DBProject.RevitServerPath3)}, {nameof(DBProject.RevitServerPath4)}) " +
                    $"VALUES (@{nameof(DBProject.Name)}, @{nameof(DBProject.Code)}, @{nameof(DBProject.Stage)}, @{nameof(DBProject.RevitVersion)}, @{nameof(DBProject.MainPath)}, " +
                        $"@{nameof(DBProject.RevitServerPath)}, @{nameof(DBProject.RevitServerPath2)}, @{nameof(DBProject.RevitServerPath3)}, @{nameof(DBProject.RevitServerPath4)}) " +
                    $"RETURNING Id;",
                    dBProject)
                .FirstOrDefault();

                return createdPrjId;
            });
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить коллекцию ВСЕХ проектов
        /// </summary>
        [Obsolete]
        public IEnumerable<DBProject> GetDBProjects() =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить коллекцию ВСЕХ проектов для нужной версии Revit
        /// </summary>
        public IEnumerable<DBProject> GetDBProjects_ByRVersion(int rVersion) =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProject.RevitVersion)}={rVersion};");

        /// <summary>
        /// Получить коллекцию ВСЕХ открытых проектов
        /// </summary>
        [Obsolete]
        public IEnumerable<DBProject> GetDBProjects_Opened() =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProject.IsClosed)}='False';");

        /// <summary>
        /// Получить коллекцию ВСЕХ открытых проектов для нужной версии Revit
        /// </summary>
        public IEnumerable<DBProject> GetDBProjects_ByRVersionANDOpened(int rVersion) =>
            ExecuteQuery<DBProject>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBProject.RevitVersion)}={rVersion} " +
                $"AND {nameof(DBProject.IsClosed)}='False';");

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
        [Obsolete]
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

        /// <summary>
        /// Получить активный проект из БД по открытому проекту Ревит
        /// </summary>
        public DBProject GetDBProject_ByRevitDocFileNameANDRVersion(string fileName, int rVersion)
        {
            IEnumerable<DBProject> filteredPrjs = GetDBProjects_ByRVersion(rVersion)
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
