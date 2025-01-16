using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using System.Collections.Generic;

namespace KPLN_IOSClasher.Services
{
    public class DBWorkerService
    {
        private readonly UserDbService _userDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;
        private readonly ProjectsIOSClashMatrixDbService _projectsIOSClashMatrixDbService;
        private readonly ProjectDbService _projectDbService;
        private DBUser _dBUser;
        private DBSubDepartment _dBSubDepartment;

        internal DBWorkerService()
        {
            // Создаю сервисы работы с БД
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
            _projectsIOSClashMatrixDbService  = (ProjectsIOSClashMatrixDbService)new CreatorProjectIOSClashMatrixDbService().CreateService();
            _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
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

        internal DBProject Get_DBProject(Document doc)
        {
            string fileName = doc.IsWorkshared
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

            return _projectDbService.GetDBProject_ByRevitDocFileName(fileName);
        }

        /// <summary>
        /// Получить коллекцию DBProjectsIOSClashMatrix, которые приняты на проекте
        /// </summary>
        /// <returns></returns>
        internal IEnumerable<DBProjectsIOSClashMatrix> Get_DBProjectsIOSClashMatrix(DBProject currentDBPrj) =>
            _projectsIOSClashMatrixDbService.GetDBProjectMatrix_ByProject(currentDBPrj);
    }
}
