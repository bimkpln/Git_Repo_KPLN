using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System;
using System.Linq;

namespace KPLN_Library_SQLiteWorker
{
    /// <summary>
    /// Сервис по обращениям к БД (криэторы лучше перевести на internal, чтобы спрятать из создание - для этого все нужно перевести на этот сервис, 
    /// сейчас созданы локально в каждом проекте)
    /// </summary>
    public static class DBMainService
    {
        private static UserDbService _userDbService;
        private static ProjectDbService _projectDbService;
        private static DocumentDbService _docDbService;
        private static ModuleDbService _moduleDbService;
        private static SubDepartmentDbService _subDepartmentDbService;
        private static ProjectsAccessMatrixDbService _projectAccessMatrixDbService;
        private static ProjectsIOSClashMatrixDbService _projectsIOSClashMatrixDbService;
        private static RevitDialogDbService _revitDialogDbService;
        private static ModuleAutostartDbService _moduleAutostartDbService;
        private static RevitDocExchangesDbService _revitDocExchangesDbService;
        private static PluginActivityDbService _pluginActivityDbService;


        private static DBUser _dBUser;
        private static DBSubDepartment _currentDBUserSubDepartment;
        private static DBSubDepartment[] _dBSubDepartmentColl;
        private static DBRevitDialog[] _dBRevitDialogColl;


        #region Создание сервисов по работе с БД
        public static UserDbService UserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

                return _userDbService;
            }
        }

        public static ProjectDbService ProjectDbService
        {
            get
            {
                if (_projectDbService == null)
                    _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();

                return _projectDbService;
            }
        }

        public static DocumentDbService DocDbService
        {
            get
            {
                if (_docDbService == null)
                    _docDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();

                return _docDbService;
            }
        }

        public static ModuleDbService ModuleDbService
        {
            get
            {
                if (_moduleDbService == null)
                    _moduleDbService = (ModuleDbService)new CreatorModuleDbService().CreateService();

                return _moduleDbService;
            }
        }

        public static SubDepartmentDbService SubDepartmentDbService
        {
            get
            {
                if (_subDepartmentDbService == null)
                    _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();

                return _subDepartmentDbService;
            }
        }

        public static ProjectsAccessMatrixDbService PrjAccessMatrixDbService
        {
            get
            {
                if (_projectAccessMatrixDbService == null)
                    _projectAccessMatrixDbService = (ProjectsAccessMatrixDbService)new CreatorProjectAccessMatrixDbService().CreateService();

                return _projectAccessMatrixDbService;
            }
        }

        public static ProjectsIOSClashMatrixDbService PrjIOSClashMatrixDbService
        {
            get
            {
                if (_projectsIOSClashMatrixDbService == null)
                    _projectsIOSClashMatrixDbService = (ProjectsIOSClashMatrixDbService)new CreatorProjectIOSClashMatrixDbService().CreateService();

                return _projectsIOSClashMatrixDbService;
            }
        }

        public static RevitDialogDbService RevitDialogDbService
        {
            get
            {
                if (_revitDialogDbService == null)
                    _revitDialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();

                return _revitDialogDbService;
            }
        }

        public static ModuleAutostartDbService ModuleAutostartDbService
        {
            get
            {
                if (_moduleAutostartDbService == null)
                    _moduleAutostartDbService = (ModuleAutostartDbService)new CreatorModuleAutostartDbService().CreateService();
                return _moduleAutostartDbService;
            }
        }

        public static RevitDocExchangesDbService RevitDocExchangesDbService
        {
            get
            {
                if (_revitDocExchangesDbService == null)
                    _revitDocExchangesDbService = (RevitDocExchangesDbService)new CreatorRevitDocExchangesDbService().CreateService();
                return _revitDocExchangesDbService;
            }
        }

        public static PluginActivityDbService PluginActivityDbService
        {
            get
            {
                if (_pluginActivityDbService == null)
                    _pluginActivityDbService = (PluginActivityDbService)new CreatorPluginActivityDbService().CreateService();
                return _pluginActivityDbService;
            }
        }
        #endregion

        #region Кэширование данных из БД
        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        public static DBUser CurrentDBUser
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
        public static DBSubDepartment CurrentUserDBSubDepartment
        {
            get
            {
                if (_currentDBUserSubDepartment == null)
                    _currentDBUserSubDepartment = SubDepartmentDbService.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _currentDBUserSubDepartment;
            }
        }

        /// <summary>
        /// Ссылка на коллецию отделов КПЛН
        /// </summary>
        public static DBSubDepartment[] DBSubDepartmentColl
        {
            get
            {
                if (_dBSubDepartmentColl == null)
                    _dBSubDepartmentColl = SubDepartmentDbService.GetDBSubDepartments().Where(dep => dep.Id != 1).ToArray();

                return _dBSubDepartmentColl;
            }
        }

        /// <summary>
        /// Список диалогов из БД
        /// </summary>
        public static DBRevitDialog[] DBRevitDialogColl
        {
            get
            {
                if (_dBRevitDialogColl == null)
                    _dBRevitDialogColl = RevitDialogDbService.GetDBRevitDialogs().ToArray();

                return _dBRevitDialogColl;
            }
        }
        #endregion

        #region Общая часть
        /// <summary>
        /// Вывод времени в определенном формате для записи в БД
        /// </summary>
        public static string CurrentTimeForDB() => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
        #endregion

        #region Дополнительные методы
        #endregion
    }
}
