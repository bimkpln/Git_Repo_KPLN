using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;

namespace KPLN_HoleManager.Common
{
    internal class DBWorkerService
    {
        private readonly UserDbService _userDbService;
        private readonly SubDepartmentDbService _subDepartmentDbService;
        private DBUser _dBUser;
        private DBSubDepartment _dBSubDepartment;

        // Создаю сервисы работы с БД
        internal DBWorkerService()
        {
            _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            _subDepartmentDbService = (SubDepartmentDbService)new CreatorSubDepartmentDbService().CreateService();
        }

        /// Ссылка на текущего пользователя из БД
        internal DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                    _dBUser = _userDbService.GetCurrentDBUser();

                return _dBUser;
            }
        }

        /// Ссылка на отдел текущего пользователя из БД
        internal DBSubDepartment CurrentDBUserSubDepartment
        {
            get
            {
                if (_dBSubDepartment == null)
                    _dBSubDepartment = _subDepartmentDbService.GetDBSubDepartment_ByDBUser(CurrentDBUser);

                return _dBSubDepartment;
            }
        }

        /// Полное имя пользователя
        internal string UserFullName
        {
            get
            {
                if (CurrentDBUser != null)
                    return $"{CurrentDBUser.Name} {CurrentDBUser.Surname}";

                System.Windows.Forms.MessageBox.Show("Не удалось получить имя пользователя из базы данных. Обратитесь в BIM-отдел.", "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

                return "ErrorName";
            }
        }

        /// Название отдела пользователя
        internal string DepartmentName
        {
            get
            {
                if (CurrentDBUserSubDepartment != null)
                    return CurrentDBUserSubDepartment.Code;

                System.Windows.Forms.MessageBox.Show("Не удалось получить отдел пользователя из базы данных. Обратитесь в BIM-отдел.", "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

                return "ErrorDepartament";
            }
        }
    }
}
