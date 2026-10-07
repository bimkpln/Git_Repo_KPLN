using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CalculateTEP.Common;
using System;

namespace KPLN_CalculateTEP.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command_AR_CalculateTEP : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            if (data.Application.ActiveUIDocument == null) return Result.Cancelled;
            if (data.Application.ActiveUIDocument.Document.IsFamilyDocument)
            {TaskDialog.Show("KPLN | Расчёт ТЭП","Откройте проект RVT. Расчёт не выполняется в редакторе семейства.");return Result.Cancelled;}
            TepCalculation.Engine.TraceNavigation("Command started: "+typeof(Command_AR_CalculateTEP).Assembly.Location);
            try
            {
                var engine = new TepCalculation.Engine(data.Application);
                do
                {
                    var window = new KPLN_CalculateTEP.Forms.AR_CalculateTEP(engine);
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = data.Application.MainWindowHandle;
                    window.ShowDialog();
                    TepCalculation.Engine.TraceNavigation("Root dialog returned; navigation="+(engine.RequestedDetail!=null));
                    if(engine.PendingContourAction==null)break;
                    engine.PerformContourAction();
                }while(true);
                engine.ScheduleRequestedNavigation(data.Application); return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex) { TepCalculation.Engine.TraceNavigation("Command failed: "+ex);message = ex.ToString(); return Result.Failed; }
            finally { TepCalculation.Engine.TraceNavigation("Command returning to host"); }
        }
    }
}
