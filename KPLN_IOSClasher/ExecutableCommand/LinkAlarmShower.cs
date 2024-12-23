using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_IOSClasher.ExecutableCommand
{
    /// <summary>
    /// Обертка для появления сообщения после завершения редактирования (OnIdling), чтобы не таким навязчивым было
    /// </summary>
    internal class LinkAlarmShower : IExecutableCommand
    {
        public Result Execute(UIApplication app)
        {
            TaskDialog td = new TaskDialog("ВНИМАНИЕ: Вы не открыли связи!")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconError,
                MainInstruction = "При моделировании линейных эл-в ВИС - необходимо подгружать ВСЕ связи ВИС, и учитывать их. Анализ на коллизии проведен НЕ в полном объеме",
                CommonButtons = TaskDialogCommonButtons.Ok,
            };

            td?.Show();

            return Result.Succeeded;
        }
    }
}
