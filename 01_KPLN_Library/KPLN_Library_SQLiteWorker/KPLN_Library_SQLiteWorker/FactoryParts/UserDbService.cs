using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts.Abstractions;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_SQLiteWorker.FactoryParts
{
    public class UserDbService : AbsDbService
    {
        internal UserDbService(string connectionString) : base(connectionString)
        {
        }

        /// <summary>
        /// Получить пользователя по имени учетки
        /// </summary>
        /// <param name="sysUserName">Имя Revit-пользователя</param>
        /// <returns>Пользователь</returns>
        public DBUser GetDBUser_ByRevitUserName(string sysUserName) => ExecuteQuery<DBUser>($"SELECT * FROM Users WHERE {nameof(DBUser.SystemName)}='{sysUserName}';").FirstOrDefault();
        
        /// <summary>
        /// Получить коллекцию ВСЕХ пользователей
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public IEnumerable<DBUser> GetDBUsers() => ExecuteQuery<DBUser>($"SELECT * FROM Users;");
    }
}
