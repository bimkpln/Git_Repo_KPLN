using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TaskManager.Services
{
    /// <summary>
    /// Сервис по кэшированию работы с БД
    /// </summary>
    public static class MainDBService
    {
        private static UserDbService _userDbService;
        private static ProjectDbService _projectDbService;
        private static SubDepartmentDbService _subDepartmentDbService;

        private static DBUser _dBUser;
        private static DBSubDepartment _currentDBSubDepartment;
        private static DBSubDepartment[] _dBSubDepartmentColl;

        internal static UserDbService UserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _userDbService;
            }
        }

        internal static ProjectDbService ProjectDbService
        {
            get
            {
                if (_projectDbService == null)
                    _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();

                return _projectDbService;
            }
        }

        internal static SubDepartmentDbService SubDepartmentDbService
        {
            get
            {
                if (_subDepartmentDbService == null)
                    _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();

                return _subDepartmentDbService;
            }
        }

        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal static DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                    _dBUser = UserDbService.GetCurrentDBUser();

                return _dBUser;
            }
        }

        /// <summary>
        /// Ссылка на отдел текущего пользователя из БД
        /// </summary>
        internal static DBSubDepartment CurrentDBUserSubDepartment
        {
            get
            {
                if (_currentDBSubDepartment == null)
                    _currentDBSubDepartment = SubDepartmentDbService.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _currentDBSubDepartment;
            }
        }

        /// <summary>
        /// Ссылка на коллецию отделов КПЛН
        /// </summary>
        internal static DBSubDepartment[] DBSubDepartmentColl
        {
            get
            {
                if (_dBSubDepartmentColl == null)
                    _dBSubDepartmentColl = SubDepartmentDbService.GetDBSubDepartments().Where(dep => dep.Id != 1).ToArray();

                return _dBSubDepartmentColl;
            }
        }

        /// <summary>
        /// Получить коллекцию специалистов отдела
        /// </summary>
        /// <param name="subDepId"></param>
        /// <returns></returns>
        internal static List<DBUser> GetDBUsers_BySubDepId(int subDepId)
        {
            if (subDepId == 0 || subDepId == -1)
                return new List<DBUser>(1){new DBUser() { Id = -2, Surname = "<Сначала выбери отдел>"} };

            List<DBUser> sortedUsersFromDB = UserDbService
                .GetDBUsers_BySubDepID(subDepId)
                .OrderBy(x => x.Surname)
                .ToList();

            sortedUsersFromDB.Insert(0, new DBUser() { Id = -2, Surname = "<Если задачу в Bitrix не отправляешь, оставь выбранным это поле>" });

            return sortedUsersFromDB;
        }
    }
}
