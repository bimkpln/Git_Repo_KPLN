using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;

namespace KPLN_Tools.Common
{
    internal static class DBWorkerService
    {
        private static readonly UserDbService _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
        private static readonly DocumentDbService _documentDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();
        private static readonly ProjectDbService _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
        private static readonly SubDepartmentDbService _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
        private static DBUser _dBUser;
        private static DBSubDepartment _dBSubDepartment;

        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal static DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                    _dBUser = _userDbService.GetCurrentDBUser();

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
                if (_dBSubDepartment == null)
                    _dBSubDepartment = _subDepartmentDbService.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _dBSubDepartment;
            }
        }

        /// <summary>
        /// Ссылка на сервис UserDbService
        /// </summary>
        public static UserDbService CurrentUserDbService { get => _userDbService; }

        /// <summary>
        /// Ссылка на сервис SubDepartmentDbService
        /// </summary>
        public static SubDepartmentDbService CurrentSubDepartmentDbService { get => _subDepartmentDbService; }

        /// <summary>
        /// Ссылка на сервис ProjectDbService
        /// </summary>
        public static ProjectDbService CurrentProjectDbService { get => _projectDbService; }
    }
}
