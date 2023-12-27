using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class SubDepartmentDbService : DbService
    {
        internal SubDepartmentDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }

        /// <summary>
        /// Получить отдел по пользователю
        /// </summary>
        /// <returns>Пользователь</returns>
        public DBSubDepartment GetCurrentUserDBSubDepartment(DBUser dbUser) =>
            ExecuteQuery<DBSubDepartment>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBSubDepartment.Id)}='{dbUser.SubDepartmentId}';")
            .FirstOrDefault();
    }
}
