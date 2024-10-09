using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Tools.Forms;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandSearchRevitUser : IExternalCommand
    {
        private static UserDbService _userDbService;
        private static SubDepartmentDbService _subDepartmentDbService;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UserSearch searchForm = new UserSearch(UserDbService
                .GetDBUsers()
                .Select(dbu => new Forms.Models.VM_UserEntity(dbu, SubDepartmentDbService.GetDBSubDepartment_ByDBUser(dbu))));
            searchForm.ShowDialog();

            return Result.Succeeded;
        }

        internal UserDbService UserDbService
        {
            get
            {
                if (_userDbService == null)
                {
                    CreatorUserDbService creatorUserDbService = new CreatorUserDbService();
                    _userDbService = (UserDbService)creatorUserDbService.CreateService();
                }
                
                return _userDbService;
            }
        }

        internal SubDepartmentDbService SubDepartmentDbService
        {
            get
            {
                if (_subDepartmentDbService == null)
                {
                    CreatorSubDepartmentDbService creatorSubDepartmentDbService = new CreatorSubDepartmentDbService();
                    _subDepartmentDbService = (SubDepartmentDbService)creatorSubDepartmentDbService.CreateService();
                }

                return _subDepartmentDbService;
            }
        }
    }
}
