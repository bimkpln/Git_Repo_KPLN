using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Parameters_Ribbon.Forms;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandSumParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (SumParameters.TryActivateExisting())
                    return Result.Succeeded;

                SumParameters mainForm = new SumParameters(commandData.Application);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(mainForm);
                mainForm.Show();

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
