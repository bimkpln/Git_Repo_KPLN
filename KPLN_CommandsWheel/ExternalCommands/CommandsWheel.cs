using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Services;

namespace KPLN_CommandsWheel.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandsWheel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return CommandWindowService.ShowCommandsWheel(commandData.Application)
                ? Result.Succeeded
                : Result.Cancelled;
        }
    }
}