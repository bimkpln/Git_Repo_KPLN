using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    /// <summary>
    /// Класс для работы с листом Users в БД
    /// </summary>
    public class UserDbService : DbService
    {
        private readonly string _sysUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last();

        internal UserDbService(string connectionString, string databaseName) : base(connectionString, databaseName)
        {
        }

        /// <summary>
        /// Получить текущего пользователя
        /// </summary>
        /// <returns>Пользователь</returns>
        public DBUser GetCurrentDBUser() => 
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.SystemName)}='{_sysUserName}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить пользователя по имени учетки
        /// </summary>
        /// <param name="sysUserName">Системное имя пользователя</param>
        /// <returns>Пользователь</returns>
        public DBUser GetDBUser_ByUserName(string sysUserName) =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.SystemName)}='{sysUserName}';")
            .FirstOrDefault();

        /// <summary>
        /// Получить коллекцию ВСЕХ пользователей
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public IEnumerable<DBUser> GetDBUsers() => 
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName};");
    }
}
