using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI;
using System;

namespace KPLN_Publication.Forms
{
    internal enum SheetBoundaryDecision
    {
        Print,
        PrintAll,
        Stop
    }

    internal static class SheetBoundaryDialog
    {
        internal static SheetBoundaryDecision Choose(SheetBoundaryIssue issue, ProgressWindow progressWindow)
        {
            var dialog = new TaskDialog("KPLN: Пакетная выдача")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                MainInstruction = "Видимые элементы выходят за границы основной надписи",
                MainContent = $"Лист {issue.Sheet.SheetNumber} — {issue.Sheet.Name}\n\n" +
                    $"Найдено элементов: {issue.ElementIds.Count}. При печати возможны смещение и обрезка содержимого.\n\n" +
                    "При остановке уже напечатанные файлы сохранятся. Если выбрано объединение PDF, " +
                    "они будут объединены. Этот и последующие листы печататься не будут.\n\n" +
                    "«Распечатать для всех» отключает эти предупреждения до конца текущего запуска. " +
                    "Список напечатанных проблемных листов будет показан в итоговом отчёте.",
                ExpandedContent = issue.Details,
                CommonButtons = TaskDialogCommonButtons.None,
                AllowCancellation = true
            };
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Распечатать");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Распечатать для всех");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Остановить печать", "Открыть этот лист после завершения выдачи готовых файлов");
            // Revit разрешает назначить кнопку по умолчанию только после её добавления.
            dialog.DefaultButton = TaskDialogResult.CommandLink3;
            // Закрытие окна и Escape также завершают пакет с сохранением готовых файлов.
            // Окно прогресса работает в отдельном потоке и находится поверх остальных окон.
            // Скрываем синхронно до открытия диалога, сохраняя текущий прогресс.
            SetProgressVisible(progressWindow, false);
            TaskDialogResult result = dialog.Show();
            if (result == TaskDialogResult.CommandLink1 || result == TaskDialogResult.CommandLink2)
                SetProgressVisible(progressWindow, true);
            if (result == TaskDialogResult.CommandLink1)
                return SheetBoundaryDecision.Print;
            if (result == TaskDialogResult.CommandLink2)
                return SheetBoundaryDecision.PrintAll;
            return SheetBoundaryDecision.Stop;
        }

        private static void SetProgressVisible(ProgressWindow window, bool visible)
        {
            if (window.Dispatcher.HasShutdownStarted || window.Dispatcher.HasShutdownFinished)
                return;

            try
            {
                window.Dispatcher.Invoke(new Action(() =>
                {
                    if (window.Dispatcher.HasShutdownStarted || window.Dispatcher.HasShutdownFinished)
                        return;
                    if (visible)
                        window.Show();
                    else
                        window.Hide();
                }));
            }
            catch (OperationCanceledException) when (window.Dispatcher.HasShutdownStarted || window.Dispatcher.HasShutdownFinished)
            {
                // Пользователь уже закрыл окно прогресса.
            }
            catch (InvalidOperationException) when (window.Dispatcher.HasShutdownStarted || window.Dispatcher.HasShutdownFinished)
            {
                // Поток окна завершился между проверкой и Invoke.
            }
        }
    }
}
