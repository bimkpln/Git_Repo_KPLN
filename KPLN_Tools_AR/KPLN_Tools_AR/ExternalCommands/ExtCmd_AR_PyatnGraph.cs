using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_AR.Forms;

namespace KPLN_Tools_AR.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_AR_PyatnGraph : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AR_PyatnGraph_Main form = new AR_PyatnGraph_Main(commandData.Application);
            
            KPLN_Library_Forms.Services.WindowHandleSearch.MainWindowHandle.SetAsOwner(form);

            form.Show();

            return Result.Cancelled;
        }
    }
}
