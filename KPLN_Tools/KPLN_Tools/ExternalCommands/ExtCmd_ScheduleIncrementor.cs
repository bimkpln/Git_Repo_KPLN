using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_ScheduleIncrementor : IExternalCommand
    {
        internal const string PluginName = "Нумерация";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            View activeView = doc.ActiveView;
            if (!(activeView is ViewSchedule viewSchedule))
            {
                MessageBox.Show("Открой нужную спецификацию для анализа", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return Result.Cancelled;
            }

            try
            {
                var model = ScheduleHelper.ReadSchedule(viewSchedule);
                var window = new ScheduleMainForm(model, viewSchedule.Name);

                // Привязка к окну ревит
                var helper = new WindowInteropHelper(window) { Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle };

                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }
    }
}
