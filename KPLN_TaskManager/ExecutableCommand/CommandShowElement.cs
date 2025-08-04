using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using KPLN_TaskManager.Common;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (uidoc == null)
                return Result.Cancelled;

            List<ElementId> elemIds = new List<ElementId>();
            foreach (string strId in _taskItemEntity.ElementIds.Split(','))
            {
                if (int.TryParse(strId, out int id))
                    elemIds.Add(new ElementId(id));
                else
                    HtmlOutput.Print($"Ошибка парсинга данных - скинь разработчику! {strId} - НЕ число", MessageType.Error);
            }

            StringBuilder sb = new StringBuilder();
            foreach(ElementId id in elemIds)
            {
                Element elem = doc.GetElement(id);
                if (elem == null)
                    sb.Append(id.ToString());
            }

            if (sb.Length > 0)
            {
                DBUser createdUser = DBMainService.UserDbService.GetDBUser_ById(_taskItemEntity.CreatedTaskUserId);

                string errorId = string.Join(", ", sb.ToString());
                HtmlOutput.Print($"Проект не содержит id из списка. Либо элементы уже удалены, либо они находятся в связи: {errorId}.\n" +
                    $"Уточни информацию у постановщика {createdUser.Name} {createdUser.Surname}, либо открой указанную в задаче модель", 
                    MessageType.Error);
            }
            else
                uidoc.Selection.SetElementIds(elemIds);

            return Result.Succeeded;
        }
    }
}
