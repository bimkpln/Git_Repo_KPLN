using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Batch.Forms;
using NLog;
using RevitServerAPILib;
using System;
using System.IO;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_Batch.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowManager : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {            
            UIApplication uiapp = commandData.Application;

            BatchManager mainForm = new BatchManager(Module.CurrentLogger, uiapp);
            if (!(bool)mainForm.ShowDialog())
                return Result.Cancelled;

            return Result.Succeeded;
        }        
    }
}
