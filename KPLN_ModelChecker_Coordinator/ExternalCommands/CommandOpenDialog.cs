extern alias revit;
using KPLN_ModelChecker_Coordinator.Forms;
using revit.Autodesk.Revit.Attributes;
using revit.Autodesk.Revit.DB;
using revit.Autodesk.Revit.UI;
using System;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_Coordinator.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandOpenDialog : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                PickProjects form = new PickProjects();
                form.Show();
                return Result.Succeeded;
            }
            catch (Exception e)
            { PrintError(e); }
            return Result.Failed;
        }
    }
}
