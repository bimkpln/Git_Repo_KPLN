﻿using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Looker.Services
{
    /// <summary>
    /// Сервис по обращению к БД
    /// </summary>
    public class DBWorkerService
    {
        private readonly UserDbService _userDbService;
        private readonly DocumentDbService _documentDbService;
        private readonly ProjectDbService _projectDbService;
        private readonly ProjectsAccessMatrixDbService _projectAccessMatrixDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;
        private readonly RevitDialogDbService _dialogDbService;
        private DBUser _dBUser;
        private DBProjectsAccessMatrix[] _dbProjectAccessMatrixColl;
        private DBSubDepartment _dBSubDepartment;
        private DBRevitDialog[] _dBRevitDialogs;

        internal DBWorkerService()
        {
            // Создаю сервисы работы с БД
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _documentDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();
            _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
            _projectAccessMatrixDbService = (ProjectsAccessMatrixDbService)new CreatorProjectAccessMatrixDbService().CreateService();
            _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
            _dialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
        }

        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal DBUser CurrentDBUser
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
        internal DBSubDepartment CurrentDBUserSubDepartment
        {
            get
            {
                if (_dBSubDepartment == null)
                    _dBSubDepartment = _subDepartmentDbService.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _dBSubDepartment;
            }
        }

        /// <summary>
        /// Ссылка на коллекцию DBProjectMatrix - матрицу допуска к проектам KPLN
        /// </summary>
        internal DBProjectsAccessMatrix[] CurrentDBProjectMatrixColl
        {
            get
            {
                if (_dbProjectAccessMatrixColl == null)
                    _dbProjectAccessMatrixColl = _projectAccessMatrixDbService.GetDBProjectMatrix().ToArray();

                return _dbProjectAccessMatrixColl;
            }
        }

        /// <summary>
        /// Список диалогов из БД
        /// </summary>
        internal DBRevitDialog[] DBRevitDialogs
        {
            get
            {
                if (_dBRevitDialogs == null)
                    _dBRevitDialogs = _dialogDbService.GetDBRevitDialogs().ToArray();

                return _dBRevitDialogs;
            }
        }


        /// <summary>
        /// Создать новый документ в БД
        /// </summary>
        internal DBDocument Create_DBDocument(string centralPath, int dBProjectId, int dBSubDepartmentId, int dbUserId, string lastChangedData, bool isClosed)
        {
            DBDocument dBDocument = new DBDocument()
            {
                CentralPath = centralPath,
                ProjectId = dBProjectId,
                SubDepartmentId = dBSubDepartmentId,
                LastChangedUserId = dbUserId,
                LastChangedData = lastChangedData,
                IsClosed = isClosed,
            };
            Task createNewDoc = Task.Run(() =>
            {
                _documentDbService.CreateDBDocument(dBDocument);
            });

            return dBDocument;
        }

        /// <summary>
        /// Получить активный проект из БД по открытому проекту Ревит
        /// </summary>
        /// <param name="fileName">Имя открытого файла Ревит</param>
        /// <returns></returns>
        internal DBProject Get_DBProjectByRevitDocFile(string fileName) => _projectDbService.GetDBProject_ByRevitDocFileName(fileName);
        

        /// <summary>
        /// Получить активный документ из БД по открытому проекту Ревит и по проекту из БД
        /// </summary>
        /// <param name="centralPath">Путь открытого файла Ревит</param>
        /// <param name="dBProject">Проект из БД</param>
        /// <returns></returns>
        internal DBDocument Get_DBDocumentByRevitDocPathAndDBProject(string centralPath, DBProject dBProject, int dBSubDepartmentId)
        {
            IEnumerable<DBDocument> docColl = _documentDbService.GetDBDocuments_ByPrjIdAndSubDepId(dBProject.Id, dBSubDepartmentId);

            return docColl.FirstOrDefault(d => d.CentralPath.Equals(centralPath));
        }

        /// <summary>
        /// Получить ID отдела, которому принадлежит файл
        /// </summary>
        /// <returns></returns>
        internal int Get_DBDocumentSubDepartmentId(Document doc)
        {
            DBSubDepartment subDep = _subDepartmentDbService.GetDBSubDepartment_ByRevitDoc(doc);
            if (subDep != null)
                return subDep.Id;

            return -1;
        }

        /// <summary>
        /// Обновление статуса документа IsClosed по статусу проекта
        /// </summary>
        /// <param name="dBProject"></param>
        internal void Update_DBDocumentIsClosedStatus(DBProject dBProject) =>
            Task.Run(() =>
            {
                _documentDbService.UpdateDBDocument_IsClosedByProject(dBProject);
            });

        /// <summary>
        /// Обновление даты последнего изменения документа
        /// </summary>
        /// <param name="dBProject"></param>
        internal void Update_DBDocumentLastChangedData(DBDocument dBDocument) =>
            Task.Run(() =>
            {
                _documentDbService.UpdateDBDocument_LastChangedData(dBDocument, CurrentDBUser, CurrentTimeForDB());
            });

        /// <summary>
        /// Вывод времени в определенном формате для записи в БД
        /// </summary>
        /// <returns></returns>
        internal string CurrentTimeForDB() => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");

        /// <summary>
        /// Метод аннулирования хранимого кэша для перезаписи при обращении
        /// </summary>
        internal void DropMainCash()
        {
            _dbProjectAccessMatrixColl = null;
            _dBRevitDialogs = null;
        }
    }
}
