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
            {TaskDialog.Show("Расчёт ТЭП","Откройте проект RVT. Расчёт не выполняется в редакторе семейства.");return Result.Cancelled;}
            try
            {
                var engine = new TepCalculation.Engine(data.Application);
                do
                {
                    var window = new KPLN_CalculateTEP.Forms.AR_CalculateTEP(engine);
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = data.Application.MainWindowHandle;
                    window.ShowDialog();
                    if(engine.PendingContourAction==null)break;
                    engine.PerformContourAction();
                }while(true);
                engine.NavigateRequested(); return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }
            catch (Exception ex) { message = ex.ToString(); return Result.Failed; }
        }
    }
}
