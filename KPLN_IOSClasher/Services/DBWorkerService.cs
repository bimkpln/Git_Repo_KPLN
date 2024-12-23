using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;

namespace KPLN_IOSClasher.Services
{
    public class DBWorkerService
    {
        private readonly UserDbService _userDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;
        private DBUser _dBUser;
        private DBSubDepartment _dBSubDepartment;

        internal DBWorkerService()
        {
            // Создаю сервисы работы с БД
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
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
    }
}
