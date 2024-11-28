using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System;
using System.Collections.Generic;
using System.Linq;

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
        /// Получить отдел по пользователю
        /// </summary>
        public IEnumerable<DBSubDepartment> GetDBSubDepartments() =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName};");

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
        public DBSubDepartment GetDBSubDepartment_ByRevitDoc(Document revitDoc)
        {
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
                    throw new Worker_Error("Не удалось определить раздел по имени файла.");
            
                return GetDBSubDepartment_SubDepartmentMatrixCode(resultSD);
            }
        }
    }
}
