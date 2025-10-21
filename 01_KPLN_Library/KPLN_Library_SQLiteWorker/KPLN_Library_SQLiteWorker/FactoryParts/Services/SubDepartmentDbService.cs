using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class SubDepartmentDbService : DbService
    {
        private readonly SubDepartmentCodeMatrixDbService _subDepartmentCodeMatrixDbService;
        
        internal SubDepartmentDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
            _subDepartmentCodeMatrixDbService = (SubDepartmentCodeMatrixDbService)new CreatorSubDepartmentCodeMatrixDbService().CreateService();
        }

        /// <summary>
        /// Получить отделы КПЛН
        /// </summary>
        public IEnumerable<DBSubDepartment> GetDBSubDepartments() =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить отдел по ID
        /// </summary>
        public DBSubDepartment GetDBSubDepartment_ById(int id) =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBSubDepartment.Id)}='{id}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить отдел по пользователю
        /// </summary>
        public DBSubDepartment GetDBSubDepartment_ByDBUser(DBUser dbUser) =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBSubDepartment.Id)}='{dbUser.SubDepartmentId}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить отдел по имени подраздела
        /// </summary>
        public DBSubDepartment GetDBSubDepartment_ByName(string name) =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBSubDepartment.Name)}='{name}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить отдел по матрице вариации кодов
        /// </summary>
        public DBSubDepartment GetDBSubDepartment_SubDepartmentMatrixCode(DBSubDepartmentCodeMatrix codeMatrix) =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBSubDepartment.Id)}='{codeMatrix.SubDepartmentId}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить отдел по файлу Revit
        /// </summary>
        [Obsolete("Лучше брать сразу по имени - так можно обработать LinkInst")]
        public DBSubDepartment GetDBSubDepartment_ByRevitDoc(Document revitDoc)
        {
            if (revitDoc == null || revitDoc.IsFamilyDocument)
                return null;
            
            string docName = revitDoc.PathName;

            if (docName.Contains("_КФ."))
                return GetDBSubDepartment_ByName("BIM");
            else
            {
                IEnumerable <DBSubDepartmentCodeMatrix> sdMatrix = _subDepartmentCodeMatrixDbService.GetSubDepartmentCodeMatrix();
                DBSubDepartmentCodeMatrix resultSD = null;
                foreach (DBSubDepartmentCodeMatrix dbSubDepartmentCodeMatrix in sdMatrix)
                {
                    if (docName.Contains($"_{dbSubDepartmentCodeMatrix.Code}")
                        || docName.Contains($"{dbSubDepartmentCodeMatrix.Code}_")
                        || docName.Contains($"-{dbSubDepartmentCodeMatrix.Code}")
                        || docName.Contains($"{dbSubDepartmentCodeMatrix.Code}-"))
                    {
                        resultSD = dbSubDepartmentCodeMatrix;
                    }
                }

                if (resultSD == null)
                    throw new Worker_Error($"Не удалось определить раздел по имени файла: {revitDoc.Title}");
            
                return GetDBSubDepartment_SubDepartmentMatrixCode(resultSD);
            }
        }

        /// <summary>
        /// Получить отдел по полному пути к файлу Revit
        /// </summary>
        public DBSubDepartment GetDBSubDepartment_ByRevitDocFullPath(string docFullPath)
        {
            if (string.IsNullOrEmpty(docFullPath))
                return null;

            if (docFullPath.Contains(".rfa") || docFullPath.Contains(".rte"))
                return null;

            Dictionary<string, int> depNameDict = new Dictionary<string, int>();
            if (docFullPath.Contains("_КФ."))
                return GetDBSubDepartment_ByName("BIM");
            else
            {
                // Т.к. имя может содержать подстроки - ищу наибольшее вол-во вхождений (может быть в пути к файлу, в имени раздел+подраздел)
                IEnumerable<DBSubDepartmentCodeMatrix> sdMatrix = _subDepartmentCodeMatrixDbService.GetSubDepartmentCodeMatrix();
                foreach (DBSubDepartmentCodeMatrix dbSubDepartmentCodeMatrix in sdMatrix)
                {
                    if (docFullPath.Contains($"_{dbSubDepartmentCodeMatrix.Code}")
                        || docFullPath.Contains($"{dbSubDepartmentCodeMatrix.Code}_")
                        || docFullPath.Contains($"-{dbSubDepartmentCodeMatrix.Code}")
                        || docFullPath.Contains($"{dbSubDepartmentCodeMatrix.Code}-"))
                    {
                        int count = 0;
                        int index = 0;
                        while ((index = docFullPath.IndexOf(dbSubDepartmentCodeMatrix.Code, index)) != -1)
                        {
                            count++;
                            index += dbSubDepartmentCodeMatrix.Code.Length; 
                        }

                        string currentDepName = GetDBSubDepartment_SubDepartmentMatrixCode(dbSubDepartmentCodeMatrix).Name;
                        if (depNameDict.TryGetValue(currentDepName, out int prevCount))
                            depNameDict[currentDepName] = prevCount + count;
                        else
                            depNameDict[currentDepName] = count;
                    }
                }

                if (depNameDict.Count == 0)
                    throw new Worker_Error($"Не удалось определить раздел по имени файла: {docFullPath}");

                int tempCount = 0;
                DBSubDepartment resultSD = null;
                foreach(var kvp in depNameDict)
                {
                    if (kvp.Value > tempCount)
                    {
                        tempCount = kvp.Value;
                        resultSD = GetDBSubDepartment_ByName(kvp.Key);
                    }
                }

                return resultSD;
            }
        }
    }
}
