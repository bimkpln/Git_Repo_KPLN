using Autodesk.Revit.DB;
using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_DBWorker.FactoryParts.SQLite
{
    /// <summary>
    /// Класс для работы с листом SubDepartmentCodeMatrix в БД
    /// </summary>
    internal class SQLiteSubDepCodeMatrixService : SQLiteService
    {
        internal SQLiteSubDepCodeMatrixService(string connectionString, DBEnumerator dbEnumerator) : base(connectionString, dbEnumerator)
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
