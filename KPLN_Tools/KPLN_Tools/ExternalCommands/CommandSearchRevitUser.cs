using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Tools.Forms;
using System.Collections.Generic;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandSearchRevitUser : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            CreatorUserDbService creatorUserDbService = new CreatorUserDbService();
            UserDbService userDbService = (UserDbService)creatorUserDbService.CreateService();
            IEnumerable<DBUser> dbUsers = userDbService.GetDBUsers();

            UserSearch searchForm = new UserSearch(dbUsers);
            searchForm.ShowDialog();

            return Result.Succeeded;
        }
    }
}
