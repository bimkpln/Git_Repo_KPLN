using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace KPLN_OpeningHoleManager.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowMainMenu : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DockablePane dockableWindow = commandData.Application.GetDockablePane(Module.PaneId);
            dockableWindow.Show();

            return Result.Succeeded;
        }
    }
}
