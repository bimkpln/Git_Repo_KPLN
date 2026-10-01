using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_RevitMcpBridge.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class BridgeStatusExtCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            // Проверка загрузки модуля
            var module = Module.Current;
            if (module == null)
            {
                message = "Модуль не загружен через KPLN_Loader.";
                return Result.Failed;
            }


            // Просмотр состояния не меняет подключение
            bool isRunning = module.Server != null;
            var dialog = new TaskDialog("KPLN: мост Codex")
            {
                MainInstruction = isRunning ? "Мост работает" : "Мост остановлен",
                MainContent = isRunning
                    ? module.Server.Url + "\nRevit " + data.Application.Application.VersionNumber +
                        "\nСессия: " + module.Server.SessionId + "\nMCP-сервер автоматически находит этот экземпляр Revit."
                    : "Подключение Codex к этому экземпляру Revit выключено.\nДля подключения нажмите «Запустить мост».",
                CommonButtons = TaskDialogCommonButtons.Close,
                DefaultButton = TaskDialogResult.Close
            };

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                isRunning ? "Остановить мост" : "Запустить мост");


            // Запуск и остановка — только по отдельному выбору пользователя
            var result = dialog.Show();
            if (result != TaskDialogResult.CommandLink1)
                return Result.Succeeded;

            if (isRunning)
            {
                module.Stop();
            }
            else
            {
                module.Start();
                if (module.Server == null)
                {
                    TaskDialog.Show("KPLN: мост Codex", module.LastError ?? "Не удалось запустить мост.");
                    return Result.Cancelled;
                }
            }


            return Result.Succeeded;
        }
    }
}
