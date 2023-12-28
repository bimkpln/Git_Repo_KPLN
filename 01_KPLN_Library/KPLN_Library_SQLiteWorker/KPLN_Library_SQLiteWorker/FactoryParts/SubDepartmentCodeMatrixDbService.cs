using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом SubDepartmentCodeMatrix в БД
    /// </summary>
    internal class SubDepartmentCodeMatrixDbService : DbService
    {
        internal SubDepartmentCodeMatrixDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }

        /// <summary>
        /// Получить варианты именования разделов
        /// </summary>
        internal IEnumerable<DBSubDepartmentCodeMatrix> GetSubDepartmentCodeMatrix() =>
            ExecuteQuery<DBSubDepartmentCodeMatrix>(
                $"SELECT * FROM {_dbTableName};");
    }
}
