using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.ExecutableCommand;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Открытие вида не может проходить в момент обработки событий, в том числе OnIdling. Подробнее тут:
    /// https://www.revitapidocs.com/2023/b6adb74b-39af-9213-c37b-f54db76b75a3.htm
    /// Поэтом - это клас для активации вида с размерами
    /// </summary>
    internal static class CheckDimension_OpenView
    {
        public static void OpenViewForDimensions(UIApplication app, Element elem)
        {
            if (!elem.IsValidObject)
            {
                TaskDialog.Show(
                    "Внимание", $"Элемент был удалён из проекта. Можно обновить список перезапустив плагин",
                    TaskDialogCommonButtons.Ok);
                return;
            }
            
            Element elemToSelec = null;
            if (elem is Dimension dim)
            {
                app.ActiveUIDocument.ActiveView = dim.View;
                elemToSelec = elem;
            }
            else if (elem is DimensionType dimType)
            {
                Document doc = app.ActiveUIDocument.Document;
                Dimension firstDim = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(el => el.GetTypeId().IntegerValue == dimType.Id.IntegerValue)
                    .Cast<Dimension>()
                    .FirstOrDefault();

                if (firstDim == null)
                {
                    TaskDialog.Show(
                        "Внимание", $"Тип размера {dimType.Name} не используется в проекте, его можно удалить, или исправить в списке типов",
                        TaskDialogCommonButtons.Ok);
                    return;
                }
                app.ActiveUIDocument.ActiveView = firstDim.View;
                elemToSelec = firstDim;
            }

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandSelectElements(new List<Element>(1) { elemToSelec }));
        }
    }
}
