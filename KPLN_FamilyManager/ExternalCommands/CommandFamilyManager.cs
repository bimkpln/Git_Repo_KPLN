using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_FamilyManager.Docking;

namespace KPLN_FamilyManager.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandFamilyManager : IExternalCommand
    {
        internal const string PluginName = "Менеджер\nсемейств";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            FamilyManagerDock.Toggle(commandData.Application);
            return Result.Succeeded;
        }

    }
}
