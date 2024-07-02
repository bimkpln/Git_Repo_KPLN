using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckLaunchDate : IExternalCommand
    {
        private static ExtensibleStorageEntity[] _extensibleStorageEntities;

        public CommandCheckLaunchDate()
        {
        }

        internal CommandCheckLaunchDate(ExtensibleStorageEntity[] extensibleStorageEntities)
        {
            _extensibleStorageEntities = extensibleStorageEntities;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            OutputFormLaunchDate form = new OutputFormLaunchDate(doc, _extensibleStorageEntities);
            form.Show();

            return Result.Succeeded;
        }
    }
}
