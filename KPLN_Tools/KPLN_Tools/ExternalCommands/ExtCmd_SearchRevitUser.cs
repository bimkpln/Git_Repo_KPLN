using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Tools.Forms;
using System.Linq;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmd_SearchRevitUser : IExternalCommand
    {
        internal const string PluginName = "Найти пользователя";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            if (uiDocument == null)
                return Result.Cancelled;

            Document document = uiDocument.Document;
            ICollection<ElementId> selectedIds = uiDocument.Selection.GetElementIds();
            Dictionary<string, List<string>> roles = null;
            string selectionInfo = null;

            // Сведения берём из того же API, что и подсказка совместной работы Revit.
            if (selectedIds.Count > 0)
            {
                roles = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                int unavailableCount = 0;
                if (document.IsWorkshared)
                {
                    foreach (ElementId id in selectedIds)
                    {
                        try
                        {
                            using (WorksharingTooltipInfo info = WorksharingUtils.GetWorksharingTooltipInfo(document, id))
                            {
                                AddRole(roles, info.Creator, "Автор");
                                AddRole(roles, info.Owner, "Текущий владелец");
                                AddRole(roles, info.LastChangedBy, "Автор последнего обновления в хранилище");
                                foreach (string requester in info.GetRequesters())
                                    AddRole(roles, requester, "Автор запроса");
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.ArgumentException)
                        {
                            unavailableCount++;
                        }
                    }
                }

                selectionInfo = $"Выбрано элементов: {selectedIds.Count}.";
                if (!document.IsWorkshared)
                    selectionInfo = "В модели отключена совместная работа: сведения об авторах недоступны.";
                else if (roles.Count == 0)
                    selectionInfo = "Для выбранных элементов нет сведений о пользователях.";

                if (unavailableCount > 0)
                    selectionInfo += $" Не удалось прочитать элементов: {unavailableCount}.";
            }

            UserSearch searchForm = new UserSearch(SQLiteMainService
                .SQLiteUserServiceInst
                .GetDBUsers()
                .Select(dbu => new Forms.Models.VM_UserEntity(dbu, SQLiteMainService.SQLiteSubDepServiceInst.GetDBSubDepartment_ByDBUser(dbu))), roles, selectionInfo, selectedIds.Count);
            searchForm.ShowDialog();

            return Result.Succeeded;
        }

        private static void AddRole(Dictionary<string, List<string>> roles, string userName, string role)
        {
            if (string.IsNullOrWhiteSpace(userName))
                return;

            userName = userName.Trim();
            string marker = userName.Trim('(', ')').Trim();
            if (string.Equals(marker, "Нет", StringComparison.OrdinalIgnoreCase)
                || string.Equals(marker, "None", StringComparison.OrdinalIgnoreCase))
                return;

            if (!roles.TryGetValue(userName, out List<string> userRoles))
            {
                userRoles = new List<string>();
                roles.Add(userName, userRoles);
            }

            if (!userRoles.Contains(role))
                userRoles.Add(role);
        }
    }
}
