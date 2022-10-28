using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.ExternalCommands;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_User.Common;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class NewCommandRun_Test : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                KPLN_ModelChecker_Lib.ExternalCommands.CommandCheckDimensions commandCheckDimensions = new KPLN_ModelChecker_Lib.ExternalCommands.CommandCheckDimensions(uiapp);
                IList<ElementEntity> errorCollection = commandCheckDimensions.Run();
                
                WPFReport wpfReport = new WPFReport(errorCollection);
                wpfReport.ShowResult();
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                PrintError(ex);
                
                return Result.Failed;
            }
        }
    }
}
