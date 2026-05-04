using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Открытие вида не может проходить в момент обработки событий, в том числе OnIdling. Подробнее тут:
    /// https://www.revitapidocs.com/2023/b6adb74b-39af-9213-c37b-f54db76b75a3.htm
    /// Поэтом - это клас для активации вида с размерами
    /// </summary>
    internal static class CheckSchedules_OpenView
    {
        public static void OpenViewForViewSchedules(UIApplication app, Element elem)
        {
            if (!elem.IsValidObject)
            {
                TaskDialog.Show(
                    "Внимание", $"Элемент был удалён из проекта. Можно обновить список перезапустив плагин",
                    TaskDialogCommonButtons.Ok);
                return;
            }

            if (elem is ViewSchedule vSch)
                app.ActiveUIDocument.ActiveView = vSch;
            else
            {
                TaskDialog.Show(
                    "Ошибка", $"Отправь разработчику - не удалось привести тип при открытии спецификации",
                    TaskDialogCommonButtons.Ok);
                return;
            }
        }
    }
}
