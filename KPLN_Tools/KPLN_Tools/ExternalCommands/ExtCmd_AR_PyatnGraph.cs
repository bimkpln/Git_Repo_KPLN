using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Forms;
using System;
using System.Windows.Interop;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_AR_PyatnGraph : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AR_PyatnGraph_Main form = new AR_PyatnGraph_Main(commandData.Application);
            
            // Связываю с окном ревит, откуда был запуск
            IntPtr windHandle = commandData.Application.MainWindowHandle;
            new WindowInteropHelper(form)
            {
                Owner = windHandle
            };

            form.Show();

            return Result.Cancelled;
        }
    }
}
