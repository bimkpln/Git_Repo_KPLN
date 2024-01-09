﻿using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Looker.Services
{
    /// <summary>
    /// Сервис по обращению к БД
    /// </summary>
    internal class DBWorkerService
    {
        private readonly UserDbService _userDbService;
        private readonly DocumentDbService _documentDbService;
        private readonly ProjectDbService _projectDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;
        private DBUser _dBUser;
        private DBSubDepartment _dBSubDepartment;

        internal DBWorkerService()
        {
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();

            _documentDbService = (DocumentDbService)new CreatorDocumentDbService().CreateService();
            _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
            _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
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
        /// Создать новый документ в БД
        /// </summary>
        /// <param name="title"></param>
        /// <param name="fileName"></param>
        /// <param name="dBProjectId"></param>
        /// <param name="dBSubDepartmentId"></param>
        /// <param name="dbUserId"></param>
        /// <param name="lastChangedData"></param>
        /// <param name="isClosed"></param>
        /// <returns></returns>
        internal DBDocument Create_DBDocument(string title, string fileName, int dBProjectId, int dBSubDepartmentId, int dbUserId, string lastChangedData, bool isClosed)
        {
            DBDocument dBDocument = new DBDocument()
            {
                Name = title,
                FullPath = fileName,
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
        internal DBProject Get_DBProjectByRevitDocFile(string fileName) => 
            _projectDbService.GetDBProjects().First(p => fileName.Contains(p.MainPath));

        /// <summary>
        /// Получить активный документ из БД по открытому проекту Ревит и по проекту из БД
        /// </summary>
        /// <param name="doc">Активынй проект Ревит</param>
        /// <param name="fileName">Имя открытого файла Ревит</param>
        /// <param name="dBProject">Проект из БД</param>
        /// <returns></returns>
        internal DBDocument Get_DBDocumentByRevitDocAndSubDepartmentAndDBProject(Document doc, string fileName, DBProject dBProject)
        {
            DBDocument dBDocument = null;
            
            DBSubDepartment dBSubDepartment = _subDepartmentDbService.GetDBSubDepartment_ByRevitDoc(doc);
            dBDocument = _documentDbService
                .GetDBDocuments_ByPrjIdAndSubDepId(dBProject.Id, dBSubDepartment.Id)
                .First(d => d.FullPath.Equals(fileName));
            if (dBDocument == null)
            {
                dBDocument = Create_DBDocument(
                    doc.Title,
                    fileName,
                    dBProject.Id,
                    dBSubDepartment.Id,
                    _dBUser.Id,
                    CurrentTimeForDB(),
                    false);
            }

            return dBDocument;
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
                _documentDbService.UpdateDBDocument_LastChangedData(dBDocument, CurrentTimeForDB());
            });

        /// <summary>
        /// Вывод времени в определенном формате для записи в БД
        /// </summary>
        /// <returns></returns>
        private string CurrentTimeForDB() => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
    }
}
