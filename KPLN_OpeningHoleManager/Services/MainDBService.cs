﻿using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Сервис по кэшированию работы с БД
    /// </summary>
    internal static class MainDBService
    {
        private static UserDbService _userDbService;
        private static ProjectDbService _projectDbService;
        private static DocumentDbService _docDbService;
        private static SubDepartmentDbService _subDepartmentDbService;
        private static RevitDialogDbService _dialogDbService;

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

        internal static DocumentDbService DocDbService
        {
            get
            {
                if (_docDbService == null)
                    _docDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();

                return _docDbService;
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

        internal static RevitDialogDbService RevitDialogDbService
        {
            get
            {
                if (_dialogDbService == null)
                    _dialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
                
                return _dialogDbService;
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
        /// Список диалогов из БД
        /// </summary>
        internal static DBRevitDialog[] DBRevitDialogs
        {
            get => RevitDialogDbService.GetDBRevitDialogs().ToArray();
        }

        /// <summary>
        /// Получить отдел, которому принадлежит файл
        /// </summary>
        /// <returns></returns>
        internal static DBSubDepartment Get_DBDocumentSubDepartment(Document doc) => SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(doc.PathName);

    }
}
