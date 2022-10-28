extern alias revit;

using KPLN_ModelChecker_Coordinator.Forms;
using revit::Autodesk.Revit.Attributes;
using revit::Autodesk.Revit.DB;
using revit::Autodesk.Revit.UI;
using System;

namespace KPLN_ModelChecker_Coordinator.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowSettings : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UserSettings form = new UserSettings();
                form.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
    }
}
