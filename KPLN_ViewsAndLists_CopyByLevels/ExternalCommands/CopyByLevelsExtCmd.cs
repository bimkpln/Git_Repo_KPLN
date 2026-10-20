using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.ViewModels;
using KPLN_ViewsAndLists_CopyByLevels.Forms;
using KPLN_ViewsAndLists_CopyByLevels.Forms.ViewModels;
using KPLN_ViewsAndLists_CopyByLevels.Services;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace KPLN_ViewsAndLists_CopyByLevels.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CopyByLevelsExtCmd : IExternalCommand
    {
        public const string PluginName = "Листы по этажам";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication app = commandData.Application;
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null || doc.IsFamilyDocument || doc.IsReadOnly) return Result.Cancelled;

            try
            {
                // Учитываем лист, выделенный до запуска команды, в том числе в диспетчере проекта.
                var selectedSheets = app.ActiveUIDocument.Selection.GetElementIds()
                    .Select(doc.GetElement).OfType<ViewSheet>()
                    .Where(sheet => !sheet.IsPlaceholder).Take(2).ToList();
                ViewSheet preselectedSheet = selectedSheets.Count == 1 ? selectedSheets[0] : null;
                var vm = new CopyByLevelsVM(doc, preselectedSheet);
                if (vm.Sheets.Count == 0)
                {
                    TaskDialog.Show("KPLN", "В проекте нет листов. Сначала подготовьте эталонный лист с планом.");
                    return Result.Cancelled;
                }
                var window = new CopyByLevelsForm(vm);
                new WindowInteropHelper(window).Owner = app.MainWindowHandle;
                if (window.ShowDialog() != true) return Result.Cancelled;


                // Модальное окно завершено: изменения выполняются в исходном API-контексте команды.
                var progress = new ProgressInfoViewModel
                {
                    ProcessTitle = PluginName,
                    MaintStatus = "Подготовка копирования листов",
                    MaxProgress = vm.Model.Preview.Count,
                    IsCancellationEnabled = false
                };
                ProgressWindow progressWindow = ProgressWindow.ShowInBackground(progress, "KPLN: " + PluginName);
                GenerationResult result;
                try
                {
                    result = GenerationService.Run(vm.Model, progress);
                    progress.IsComplete = true;
                }
                finally
                {
                    progressWindow.CloseSafely();
                }

                if (result.FirstSheet != null)
                {
                    try { app.ActiveUIDocument.ActiveView = (View)doc.GetElement(result.FirstSheet); }
                    catch (Exception) { /* Созданные листы доступны в диспетчере проекта. */ }
                }
                ShowReport(app, result);
                // Ошибки отдельных этажей уже откатились; Failed здесь отменил бы и успешные этажи.
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                TaskDialog.Show("KPLN: " + PluginName, ex.Message);
                return Result.Failed;
            }
        }

        private static void ShowReport(UIApplication app, GenerationResult result)
        {
            string summary = "Создано: " + result.Created + "   Пропущено: " + result.Skipped + "   Ошибок: " + result.Failed;
            var window = new KPLN_Library_Forms.UI.OperationReportWindow(
                "KPLN: Результат создания листов", result.Report, summary, result.ReportPath);
            new WindowInteropHelper(window).Owner = app.MainWindowHandle;
            window.ShowDialog();
        }
    }
}
