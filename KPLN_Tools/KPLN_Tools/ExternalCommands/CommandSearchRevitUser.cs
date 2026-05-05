using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Library_PluginActivityWorker;
using KPLN_Tools.Forms;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandSearchRevitUser : IExternalCommand
    {
        internal const string PluginName = "Найти пользователя";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);
            
            UserSearch searchForm = new UserSearch(SQLiteMainService
                .SQLiteUserServiceInst
                .GetDBUsers()
                .Select(dbu => new Forms.Models.VM_UserEntity(dbu, SQLiteMainService.SQLiteSubDepServiceInst.GetDBSubDepartment_ByDBUser(dbu))));
            searchForm.ShowDialog();

            return Result.Succeeded;
        }
    }
}
