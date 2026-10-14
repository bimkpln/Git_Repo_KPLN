using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UI.ViewModels;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
using System.Diagnostics;
using System.Windows;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class OVVK_SystemRefreshExcCmd : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly Window _owner;

        internal OVVK_SystemRefreshExcCmd(OVVK_SystemManagerM model, Window owner)
        {
            _model = model;
            _owner = owner;
        }

        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null || !doc.Equals(_model.CurrentDoc))
                return Result.Cancelled;
            if (doc.IsReadOnly)
                throw new InvalidOperationException("Документ доступен только для чтения.");

            var progress = new ProgressInfoViewModel
            {
                ProcessTitle = "Обновление имени системы",
                MaintStatus = "Собираю системы модели",
                IsIndeterminate = true,
            };
            var window = ProgressWindow.ShowInBackground(progress, "KPLN: Обновление систем");
            var service = new SystemRefreshService(doc);
            try
            {
                Action checkCancellation = () =>
                {
                    if (progress.IsCancellationRequested)
                        throw new OperationCanceledException();
                };
                var clock = Stopwatch.StartNew();
                string lastStage = null;
                service.Execute((completed, total, stage) =>
                {
                    checkCancellation();
                    if (lastStage == stage && completed < total && clock.ElapsedMilliseconds < 75)
                        return;

                    clock.Restart();
                    lastStage = stage;
                    progress.MaintStatus = stage;
                    progress.MaxProgress = Math.Max(1, total);
                    progress.CurrentProgress = completed;
                    progress.IsIndeterminate = false;
                }, checkCancellation, () =>
                {
                    progress.IsCancellationEnabled = false;
                    checkCancellation();
                    progress.MaintStatus = "Сохраняю изменения";
                    progress.IsIndeterminate = true;
                });
                progress.IsComplete = true;
            }
            catch (OperationCanceledException)
            {
                _model.SetOperationMessage("Обновление систем отменено. Изменения не сохранены.");
                return Result.Cancelled;
            }
            finally
            {
                window.CloseSafely();
            }

            _model.Refresh();
            foreach (string warning in service.Warnings)
                HtmlOutput.Print(warning, MessageType.Warning);
            MessageBox.Show(_owner, $"Обновлено сокращений типов: {service.ChangedTypeCount}\n"
                + $"Обновлено систем: {service.RefreshedSystemCount}\nПропущено с замечаниями: {service.Warnings.Count}",
                "KPLN: Обновление имени системы", MessageBoxButton.OK,
                service.Warnings.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            return Result.Succeeded;
        }
    }
}
