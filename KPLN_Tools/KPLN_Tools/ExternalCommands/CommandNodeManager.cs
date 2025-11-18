using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Docking;
using KPLN_Tools.Forms;


namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandNodeManager : IExternalCommand
    {
        internal const string PluginName = "Менеджер\nузлов";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;

            var window = new MainWindowNodeManager(uiapp, uidoc);
            window.Show(); 

            return Result.Succeeded;
        }
    }
}