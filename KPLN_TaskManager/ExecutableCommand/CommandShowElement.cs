using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_TaskManager.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TaskManager.ExecutableCommand
{
    internal class CommandShowElement : IExecutableCommand
    {
        private readonly TaskItemEntity _taskItemEntity;

        public CommandShowElement(TaskItemEntity taskItemEntity)
        {
            _taskItemEntity = taskItemEntity;
        }

        public Result Execute(UIApplication app)
        {
            List<ElementId> elemIds = new List<ElementId>();
            foreach (string strId in _taskItemEntity.ElementIds.Split(','))
            {
                if (int.TryParse(strId, out int id))
                    elemIds.Add(new ElementId(id));
                else
                    HtmlOutput.Print($"Ошибка парсинга данных - скинь разработчику! {strId} - НЕ число", MessageType.Error);
            }

            app.ActiveUIDocument.Selection.SetElementIds(elemIds);

            return Result.Succeeded;
        }
    }
}
