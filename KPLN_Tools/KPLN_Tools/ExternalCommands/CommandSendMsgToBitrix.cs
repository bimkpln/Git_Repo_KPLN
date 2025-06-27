using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_PluginActivityWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using KPLN_Tools.Forms.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandSendMsgToBitrix : IExternalCommand
    {
        internal const string PluginName = "Отправить\nв Bitrix";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document selectedDoc = null;
            string selectedDocTitle = string.Empty;
            string selectedDocPath = string.Empty;
            string selectedDocActiveViewName = string.Empty;
            List<string> selectedIds = new List<string>();

            #region Анализ выборки пользователем
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection selection = uidoc.Selection;

            bool checkLinkElem = false;
#if Revit2020 || Debug2020
            // Для 2020 - нет возможности получить элементы из связи. Это исправили в более поздних версиях
            ICollection<ElementId> selIds = selection.GetElementIds();
            foreach (ElementId elementId in selIds)
            {
                Element elem = doc.GetElement(elementId);
                if (elem is RevitLinkInstance rli)
                    checkLinkElem = true;
                else
                    selectedIds.Add(elementId.ToString());
            }
            selectedDoc = doc;
#else
            ElementId elementId = null;
            bool checkDoubleLinkElem = false;
            bool checkDocElem = false;
            IList<Reference> selRefers = selection.GetReferences();
            foreach (Reference selRefer in selRefers)
            {
                Element selElem = doc.GetElement(selRefer);
                if (selElem is RevitLinkInstance rli)
                {
                    Document linkedDoc = rli.GetLinkDocument();
                    if (selectedDoc != null && selectedDoc.Title != linkedDoc.Title)
                        checkDoubleLinkElem = true;

                    selectedDoc = linkedDoc;
                    selectedDocActiveViewName = "Отправлено из связи, вид не имеет значения";

                    Element linkId = linkedDoc.GetElement(selRefer.LinkedElementId);
                    if (linkId != null)
                        elementId = linkedDoc.GetElement(selRefer.LinkedElementId).Id;
                    else
                        checkLinkElem = true;
                }
                else
                {
                    elementId = selElem.Id;
                    selectedDoc = doc;

                    Autodesk.Revit.DB.View activeView = doc.ActiveView;
                    if (activeView == null)
                        selectedDocActiveViewName = "Не удалось определить активный вид";
                    else if (activeView is ViewSheet vsh)
                        selectedDocActiveViewName = $"Лист: {vsh.SheetNumber} - {vsh.Name}";
                    else
                        selectedDocActiveViewName = $"Вид: {doc.ActiveView.Name}";

                    checkDocElem = true;
                }

                if (elementId != null)
                    selectedIds.Add(elementId.ToString());
            }


            // Проверка на выборки из разных документов
            if (checkLinkElem && checkDocElem)
            {
                MessageBox.Show($"Выбраны элементы внутри проекта, и из связи. Такой формат НЕ поддеривается, делай выборку внутри одного документа.", "KPLN", MessageBoxButtons.OK);

                return Result.Cancelled;
            }

            // Проверка на выборки из разных документов
            if (checkDoubleLinkElem)
            {
                MessageBox.Show($"Выбраны элементы внутри разных связей. Такой формат НЕ поддеривается, делай выборку внутри одного документа.", "KPLN", MessageBoxButtons.OK);

                return Result.Cancelled;
            }
#endif
            // Проверка на выборку экземпляра связи
            if (checkLinkElem)
            {
                MessageBox.Show($"Нет смысла передавать связь, выбери либо элемент из проекта, либо элемент из связи", "KPLN", MessageBoxButtons.OK);

                return Result.Cancelled;
            }

            if (selectedIds.Count == 0)
                throw new System.Exception("Ошибка получения списка ID элементов. Отправь разработчику!");

            // Настраиваю имя и путь к проекту (юзерфрендли)
            if (selectedDoc.IsWorkshared)
            {
                ModelPath selectedDocModelPath = selectedDoc.GetWorksharingCentralModelPath();
                // Обработка РС
                if (selectedDocModelPath.ServerPath)
                {
                    string centralDocPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(selectedDoc.GetWorksharingCentralModelPath());
                    string[] centralDocPathParts = centralDocPath.Split('/');

                    selectedDocPath = $"\n\t[i]Адрес Revit-Server[/i]: http://{centralDocPathParts[2]}/RevitServerAdmin{uiapp.Application.VersionNumber}\n\t[i]Путь по структуре: [/i]"
                        + string.Join("/", centralDocPathParts.Where(str => !str.Contains(".rvt")));
                    selectedDocTitle = centralDocPathParts.FirstOrDefault(str => str.Contains(".rvt"));
                }
                // Остлаьные файлы (подразумевается сервер КПЛН)
                else
                {
                    string centralDocPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(selectedDoc.GetWorksharingCentralModelPath());
                    string[] centralDocPathParts = centralDocPath.Split('\\');

                    selectedDocPath = string.Join("\\", centralDocPathParts.Where(str => !str.Contains(".rvt")));
                    selectedDocTitle = centralDocPathParts.FirstOrDefault(str => str.Contains(".rvt"));
                }

                // У отсоединенных файлов путь МХ не выцепить - они в оперативке висят
                if (string.IsNullOrEmpty(selectedDocTitle))
                    selectedDocTitle = doc.Title;
            }
            else
            {
                selectedDocPath = selectedDoc.PathName;
                selectedDocTitle = selectedDoc.Title;
            }
            #endregion

            // Обработка данных по пользователям
            IEnumerable<DBUser> dbUsers = DBWorkerService.CurrentUserDbService.GetDBUsers();

            // Обработка данных по отделам
            IEnumerable<DBSubDepartment> dbSubDeps = DBWorkerService.CurrentSubDepartmentDbService.GetDBSubDepartments();

            #region Подготовка и формирование окна
            ObservableCollection<SendMsgToBitrix_UserEntity> modelsForForm = new ObservableCollection<SendMsgToBitrix_UserEntity>(dbUsers
                .Select(user => new SendMsgToBitrix_UserEntity(user, DBWorkerService.CurrentSubDepartmentDbService.GetDBSubDepartment_ByDBUser(user)))
                .OrderByDescending(user => user.DBSubDepartment.Id)
                .ThenBy(user => user.DBUser.Id));

            SendMsgToBitrix form = new SendMsgToBitrix(modelsForForm);
            form.ShowDialog();
            #endregion

            #region Обработка результата
            if ((bool)form.DialogResult)
            {
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                string elemIds = string.Join(",", selectedIds);

                form.CurrentViewModel.MessageToSend_MainData = $"[u]Имя файла:[/u] {selectedDocTitle}\n" +
                    $"[u]Путь к файлу:[/u] {selectedDocPath}\n" +
                    $"[u]Вид, с которого отправлено:[/u] {selectedDocActiveViewName}\n" +
                    $"[u]ID элемента/-ов:[/u] {elemIds}";

                IEnumerable<SendMsgToBitrix_UserEntity> selectedUsers = form.CurrentViewModel.SelectedElements;
                foreach (SendMsgToBitrix_UserEntity entity in selectedUsers)
                {
                    string msg = $"Данные от [b]{DBWorkerService.CurrentDBUser.Surname} {DBWorkerService.CurrentDBUser.Name}[/b] из отдела {DBWorkerService.CurrentDBUserSubDepartment.Code}\n\n" +
                        $"[b]Данные по элементу:[/b]\n{form.CurrentViewModel.MessageToSend_MainData}\n\n";
                    if (!string.IsNullOrEmpty(form.CurrentViewModel.MessageToSend_UserComment))
                        msg += $"[b]Комментарий:[/b]\n {form.CurrentViewModel.MessageToSend_UserComment}";

                    BitrixMessageSender.SendMsg_ToUser_ByDBUser(entity.DBUser, msg);
                }

                MessageBox.Show($"Сообщение успешно отправлено! Скоро с вами свяжется выбранный специалист, ожидайте...", "KPLN", MessageBoxButtons.OK);

                return Result.Succeeded;
            }
            #endregion

            return Result.Cancelled;
        }
    }
}
