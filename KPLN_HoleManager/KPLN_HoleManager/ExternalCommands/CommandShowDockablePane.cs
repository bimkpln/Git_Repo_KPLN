using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_HoleManager.Common;
using System;

namespace KPLN_HoleManager.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowDockablePane : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!commandData.Application.ActiveUIDocument.Document.IsFamilyDocument)
                commandData.Application.GetDockablePane(new DockablePaneId(DockablePreferences.PageGuid)).Show();
            
            return Result.Succeeded;
        }
    }
}
