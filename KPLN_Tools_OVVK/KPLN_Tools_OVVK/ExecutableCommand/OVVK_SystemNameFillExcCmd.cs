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
using System.Linq;
using System.Windows;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class OVVK_SystemNameFillExcCmd : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly string _parameterName;
        private readonly ElementId _viewId;
        private readonly Window _owner;

        internal OVVK_SystemNameFillExcCmd(OVVK_SystemManagerM model, string parameterName, ElementId viewId, Window owner)
        {
            _model = model;
            _parameterName = parameterName;
            _viewId = viewId;
            _owner = owner;
        }


        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument?.Document;
            if (doc == null || !doc.Equals(_model.CurrentDoc) || string.IsNullOrWhiteSpace(_parameterName))
                return Result.Cancelled;
            if (doc.IsReadOnly)
                throw new InvalidOperationException("Документ доступен только для чтения.");

            // Повторная проверка перед записью: настройки типов могли измениться после открытия менеджера.
            if (!_model.ValidateSystemAbbreviations())
                return Result.Cancelled;

            var progress = new ProgressInfoViewModel
            {
                ProcessTitle = $"Заполнение «{_parameterName}»",
                MaintStatus = "Собираю элементы",
                IsIndeterminate = true,
            };
            ProgressWindow progressWindow = ProgressWindow.ShowInBackground(progress,
                "KPLN: Заполнение параметра систем");
            SystemNameFillService service;

            try
            {
                var progressClock = Stopwatch.StartNew();
                string lastStage = null;
                Action checkCancellation = () =>
                {
                    if (progress.IsCancellationRequested)
                        throw new OperationCanceledException();
                };
                Action<int, int, string> reportProgress = (completed, total, stage) =>
                {
                    checkCancellation();
                    // Частые записи параметров не должны переполнять очередь отдельного UI-потока.
                    if (stage == lastStage && completed < total && progressClock.ElapsedMilliseconds < 75)
                        return;

                    lastStage = stage;
                    progressClock.Restart();
                    progress.MaintStatus = stage;
                    progress.MaxProgress = Math.Max(1, total);
                    progress.CurrentProgress = completed;
                    progress.IsIndeterminate = false;
                };

                Element[] elements = OVVK_SystemManagerM.CollectElements(doc, _viewId,
                    (completed, total) => reportProgress(completed, total, "Собираю элементы"));
                if (elements.Length == 0)
                {
                    _model.SetOperationMessage(_viewId == null
                        ? "В модели нет элементов, соответствующих фильтру менеджера."
                        : "На текущем виде нет элементов, соответствующих фильтру менеджера.");
                    return Result.Cancelled;
                }

                service = new SystemNameFillService(doc, _parameterName, elements, checkCancellation);
                service.Prepare(reportProgress);


                using (var transaction = new Transaction(doc, "KPLN: Заполнение параметра систем"))
                {
                    transaction.Start();
                    service.Write(reportProgress);
                    progress.IsCancellationEnabled = false;
                    checkCancellation();
                    progress.MaintStatus = "Сохраняю изменения";
                    progress.IsIndeterminate = true;
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Заполнение параметра не сохранено: транзакция отменена.");
                }

                progress.IsComplete = true;
            }
            catch (OperationCanceledException)
            {
                // Незавершённая транзакция откатывается при выходе из using.
                _model.SetOperationMessage("Заполнение отменено. Изменения не сохранены.");
                return Result.Cancelled;
            }
            finally
            {
                progressWindow.CloseSafely();
            }


            _model.Refresh();
            foreach (var group in service.Errors.GroupBy(error => error.Value))
                HtmlOutput.Print(group.Key + " Элементы: " + string.Join(", ", group.Select(error => error.Key)), MessageType.Warning);
            if (!string.IsNullOrEmpty(service.LogWarning))
                HtmlOutput.Print(service.LogWarning, MessageType.Warning);

            string message = $"Параметр: {_parameterName}\nИзменено элементов: {service.ChangedCount}\n"
                + $"Уже заполнено: {service.UnchangedCount}\nПропущено с замечаниями: {service.Errors.Count}";
            if (service.TypeParameterCount > 0)
                message += $"\nИзменено параметров типа: {service.TypeParameterCount}. Значения общие для всех экземпляров этих типов.";
            bool hasWarnings = service.Errors.Count > 0 || !string.IsNullOrEmpty(service.LogWarning);
            if (hasWarnings)
                message += "\nПодробности выведены в отчёт.";

            MessageBox.Show(_owner, message, "KPLN: Заполнение параметра систем", MessageBoxButton.OK,
                hasWarnings ? MessageBoxImage.Warning : MessageBoxImage.Information);
            return Result.Succeeded;
        }
    }
}
