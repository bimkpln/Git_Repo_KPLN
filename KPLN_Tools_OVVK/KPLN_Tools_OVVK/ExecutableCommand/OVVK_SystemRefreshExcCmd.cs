using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_Forms.UI.ViewModels;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
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

            string refreshToken = _model.BeginSystemNameRefresh();

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
                service.Execute(progress.Report, progress.ThrowIfCancellationRequested, progress.BeginCommit);
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

            _model.CompleteSystemNameRefresh(refreshToken, service.Warnings.Count == 0);
            foreach (string warning in service.Warnings)
                HtmlOutput.Print(System.Net.WebUtility.HtmlEncode(warning).Replace("\n", "<br>"), MessageType.Warning);
            MessageBox.Show(_owner, $"Обновлено сокращений типов: {service.ChangedTypeCount}\n"
                + $"Обновлено систем: {service.RefreshedSystemCount}\nПропущено с замечаниями: {service.Warnings.Count}",
                "KPLN: Обновление имени системы", MessageBoxButton.OK,
                service.Warnings.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            return Result.Succeeded;
        }
    }
}
