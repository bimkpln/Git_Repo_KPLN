using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;

namespace KPLN_IOSClasher.ExecutableCommand
{
    /// <summary>
    /// Обертка для появления сообщения после завершения редактирования (OnIdling), чтобы не таким навязчивым было
    /// </summary>
    internal sealed class LinkAlarmShower : IExecutableCommand
    {
        private static DateTime _lastAlarm = new DateTime(2025, 01, 01, 0, 0, 0);
        private readonly TimeSpan _delay = new TimeSpan(0, 10, 0);

        public Result Execute(UIApplication app)
        {
            // Такое возможно при работе плагинов с открываением/сохранением моделей (модель не открылась)
            if (app.ActiveUIDocument == null)
                return Result.Cancelled;

            DateTime temp = DateTime.Now;
            if (_delay < (temp - _lastAlarm))
            {
                TaskDialog td = new TaskDialog("ВНИМАНИЕ: Вы не открыли связи!")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconError,
                    MainInstruction = "При моделировании линейных эл-в ВИС - необходимо подгружать ВСЕ связи ВИС, и учитывать их. Анализ на коллизии проведен НЕ в полном объеме",
                };

                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Сейчас загружу");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Отложить на 10 минут");

                TaskDialogResult result = td.Show();
                if (result == TaskDialogResult.CommandLink2)
                    _lastAlarm = DateTime.Now;
            }

            return Result.Succeeded;
        }
    }
}
