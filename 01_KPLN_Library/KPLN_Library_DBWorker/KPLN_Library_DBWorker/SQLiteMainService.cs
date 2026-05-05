using KPLN_Library_DBWorker.Core;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
using System;
using System.Linq;

namespace KPLN_Library_DBWorker
{
    /// <summary>
    /// Сервис по обращениям к БД (криэторы лучше перевести на internal, чтобы спрятать из создание - для этого все нужно перевести на этот сервис, 
    /// сейчас созданы локально в каждом проекте)
    /// </summary>
    public static class SQLiteMainService
    {
        private static SQLiteUserService _userService;
        private static SQLitePrjService _projectService;
        private static SQLiteDocService _docService;
        private static SQLiteModuleService _moduleService;
        private static SQLiteSubDepService _subDepartmentService;
        private static SQLitePrjAccessMatrixService _projectAccessMatrixService;
        private static SQLitePrjOSClashMatrixService _projectsIOSClashMatrixService;
        private static SQLiteRevitDialogService _revitDialogService;
        private static SQLiteModuleAutostartService _moduleAutostartService;
        private static SQLiteRevitDocExchangesService _revitDocExchangesService;
        private static SQLitePluginActivityService _pluginActivityService;


        private static DBUser _dBUser;
        private static DBSubDepartment _dbuserSubDepartment;
        private static DBSubDepartment[] _dBSubDepartmentColl;
        private static DBRevitDialog[] _dBRevitDialogColl;


        #region Создание сервисов по работе с БД
        public static SQLiteUserService SQLiteUserServiceInst
        {
            get
            {
                if (_userService == null)
                    _userService = (SQLiteUserService)new SQLiteCreatorUserDbService().CreateService();

                return _userService;
            }
        }

        public static SQLitePrjService SQLitePrjServiceInst
        {
            get
            {
                if (_projectService == null)
                    _projectService = (SQLitePrjService)new SQLiteCreatorPrjService().CreateService();

                return _projectService;
            }
        }

        public static SQLiteDocService SQLiteDocServiceInst
        {
            get
            {
                if (_docService == null)
                    _docService = (SQLiteDocService)new SQLiteCreatorDocService().CreateService();

                return _docService;
            }
        }

        public static SQLiteModuleService SQLiteModuleServiceInst
        {
            get
            {
                if (_moduleService == null)
                    _moduleService = (SQLiteModuleService)new SQLiteCreatorModuleService().CreateService();

                return _moduleService;
            }
        }

        public static SQLiteSubDepService SQLiteSubDepServiceInst
        {
            get
            {
                if (_subDepartmentService == null)
                    _subDepartmentService = (SQLiteSubDepService)new SQLiteCreatorSubDepService().CreateService();

                return _subDepartmentService;
            }
        }

        public static SQLitePrjAccessMatrixService SQLitePrjAccessMatrixServiceInst
        {
            get
            {
                if (_projectAccessMatrixService == null)
                    _projectAccessMatrixService = (SQLitePrjAccessMatrixService)new SQLiteCreatorPrjAccessMatrixService().CreateService();

                return _projectAccessMatrixService;
            }
        }

        public static SQLitePrjOSClashMatrixService SQLitePrjOSClashMatrixServiceInst
        {
            get
            {
                if (_projectsIOSClashMatrixService == null)
                    _projectsIOSClashMatrixService = (SQLitePrjOSClashMatrixService)new SQLiteCreatorPrjIOSClashMatrixService().CreateService();

                return _projectsIOSClashMatrixService;
            }
        }

        public static SQLiteRevitDialogService SQLiteRevitDialogServiceInst
        {
            get
            {
                if (_revitDialogService == null)
                    _revitDialogService = (SQLiteRevitDialogService)new SQLiteCreatorRevitDialogtService().CreateService();

                return _revitDialogService;
            }
        }

        public static SQLiteModuleAutostartService SQLiteModuleAutostartServiceInst
        {
            get
            {
                if (_moduleAutostartService == null)
                    _moduleAutostartService = (SQLiteModuleAutostartService)new SQLiteCreatorModuleAutostartService().CreateService();
                return _moduleAutostartService;
            }
        }

        public static SQLiteRevitDocExchangesService SQLiteRevitDocExchangesServiceInst
        {
            get
            {
                if (_revitDocExchangesService == null)
                    _revitDocExchangesService = (SQLiteRevitDocExchangesService)new SQLiteCreatorRevitDocExchangesService().CreateService();
                return _revitDocExchangesService;
            }
        }

        public static SQLitePluginActivityService SQLitePluginActivityServiceInst
        {
            get
            {
                if (_pluginActivityService == null)
                    _pluginActivityService = (SQLitePluginActivityService)new SQLiteCreatorPluginActivityService().CreateService();
                return _pluginActivityService;
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
                    _dBUser = SQLiteUserServiceInst.GetCurrentDBUser();

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
                if (_dbuserSubDepartment == null)
                    _dbuserSubDepartment = SQLiteSubDepServiceInst.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _dbuserSubDepartment;
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
                    _dBSubDepartmentColl = SQLiteSubDepServiceInst.GetDBSubDepartments().Where(dep => dep.Id != 1).ToArray();

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
                    _dBRevitDialogColl = SQLiteRevitDialogServiceInst.GetDBRevitDialogs().ToArray();

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
