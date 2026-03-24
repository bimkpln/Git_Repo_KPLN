using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandMovingElementsInLevel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            FormChangeElementInLevel form = new FormChangeElementInLevel(commandData.Application.ActiveUIDocument.Document);
            form.ShowDialog();

            return Result.Succeeded;
        }
    }
}