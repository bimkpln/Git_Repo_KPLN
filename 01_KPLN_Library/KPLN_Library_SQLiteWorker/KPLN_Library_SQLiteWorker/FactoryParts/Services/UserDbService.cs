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

        internal UserDbService(string connectionString, DB_Enumerator dbEnumerator) : base(connectionString, dbEnumerator)
        {
        }

        #region Create
        #endregion

        #region Read
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
        /// Получить коллекцию ВСЕХ пользователей
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public IEnumerable<DBUser> GetDBUsers() =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName};");

        /// <summary>
        /// Получить пользователя бим-отдела - руководительс
        /// </summary>
        public IEnumerable<DBUser> GetDBUsers_BIMManager() =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.SystemName)}='tkutsko';");

        /// <summary>
        /// Получить коллекцию пользователей бим-отдела - отдела АР
        /// </summary>
        public IEnumerable<DBUser> GetDBUsers_BIMARCoord() =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.SystemName)}='ukolomiec' " +
                $"OR {nameof(DBUser.SystemName)}='gfedoseeva';");

        /// <summary>
        /// Получить пользователя по Id
        /// </summary>
        /// <param name="id">Id пользователя</param>
        /// <returns>Пользователь</returns>
        public DBUser GetDBUser_ById(int id) =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.Id)}='{id}';")
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
        /// Получить коллекцию ВСЕХ пользователей ОТДЕЛА (по ID)
        /// </summary>
        /// <returns>Коллекция пользователей</returns>
        public IEnumerable<DBUser> GetDBUsers_BySubDepID(int subDepId) =>
            ExecuteQuery<DBUser>(
                $"SELECT * FROM {_dbTableName} " +
                $"WHERE {nameof(DBUser.SubDepartmentId)}='{subDepId}';");
        #endregion

        #region Update
        /// <summary>
        /// Обновить значение BitrixID для пользователя
        /// </summary>
        public void UpdateDBUser_BitrixUserID(DBUser dbUser, int bitrixId) =>
            ExecuteNonQuery($"UPDATE {_dbTableName} " +
                  $"SET {nameof(DBUser.BitrixUserID)}='{bitrixId}' WHERE {nameof(DBUser.Id)}='{dbUser.Id}';");
        #endregion

        #region Delete
        #endregion
    }
}
