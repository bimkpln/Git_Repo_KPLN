using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_DefaultPanelExtension_Modify.Forms;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.Services;
using KPLN_Library_PluginActivityWorker;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_DefaultPanelExtension_Modify.ExecutableCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class ExtCmdSendToBitrix : IExecutableCommand
    {
        internal const string PluginName = "Отправить в Bitrix";

        public ExtCmdSendToBitrix()
        {
        }

        public Result Execute(UIApplication app)
        {
#if !Debug2020 && !Revit2020
            Document selectedDoc = null;
            string selectedDocTitle = string.Empty;
            string selectedDocPath = string.Empty;
            string selectedDocActiveViewName = string.Empty;
            List<string> selectedIds = new List<string>();

            #region Анализ выборки пользователем
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection selection = uidoc.Selection;

            bool checkLinkElem = false;
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

                    selectedDocPath = $"\n\t[i]Адрес Revit-Server[/i]: http://{centralDocPathParts[2]}/RevitServerAdmin{app.Application.VersionNumber}\n\t[i]Путь по структуре: [/i]"
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
            IEnumerable<DBUser> dbUsers = DBMainService.UserDbService.GetDBUsers()
                .OrderByDescending(user => user.SubDepartmentId)
                .ThenBy(user => user.Surname);

            // Обработка данных по отделам
            IEnumerable<DBSubDepartment> dbSubDeps = DBMainService.SubDepartmentDbService.GetDBSubDepartments();

            #region Подготовка и формирование окна
            ObservableCollection<SendMsgToBitrix_UserEntity> modelsForForm = new ObservableCollection<SendMsgToBitrix_UserEntity>(dbUsers
                .Select(user => new SendMsgToBitrix_UserEntity(user, DBMainService.DBSubDepartmentColl.FirstOrDefault(sd => sd.Id == user.SubDepartmentId))));

            SendMsgToBitrix form = new SendMsgToBitrix(modelsForForm);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(form);
            form.ShowDialog();
            #endregion

            #region Обработка результата
            if ((bool)form.DialogResult)
            {
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(PluginName, ModuleData.ModuleName).ConfigureAwait(false);

                string elemIds = string.Join(",", selectedIds);

                form.CurrentViewModel.MessageToSend_MainData = $"[u]Имя файла:[/u] {selectedDocTitle}\n" +
                    $"[u]Путь к файлу:[/u] {selectedDocPath}\n" +
                    $"[u]Вид, с которого отправлено:[/u] \"{selectedDocActiveViewName}\"\n" +
                    $"[u]ID элемента/-ов:[/u] {elemIds}";

                IEnumerable<SendMsgToBitrix_UserEntity> selectedUsers = form.CurrentViewModel.SelectedElements;
                string msg = $"Данные от [b]{DBMainService.CurrentDBUser.Surname} {DBMainService.CurrentDBUser.Name}[/b] из отдела {DBMainService.CurrentUserDBSubDepartment.Code}\n\n" +
                    $"[b]Данные по элементу:[/b]\n{form.CurrentViewModel.MessageToSend_MainData}\n\n";
                if (!string.IsNullOrEmpty(form.CurrentViewModel.MessageToSend_UserComment))
                    msg += $"[b]Комментарий:[/b]\n {form.CurrentViewModel.MessageToSend_UserComment}";

                // Отправляю сообщение в чат
                string imgId = string.Empty;
                if (form.CurrentViewModel.MsgImageSource != null)
                {
                    // Гружу картинки в битрикс
                    Task<string> bitrUploadImageIdTask = Task<string>.Run(() =>
                    {
                        return BitrixMessageSender.UploadFile_ToSpecialFolder(form.CurrentViewModel.ImageBuffer, $"BMS_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}");
                    });
                    bitrUploadImageIdTask.Wait();

                    imgId = bitrUploadImageIdTask?.Result;
                    if (string.IsNullOrWhiteSpace(imgId))
                    {
                        MessageBox.Show($"Ошбка отправки сообщения :(", "KPLN", MessageBoxButtons.OK);
                        return Result.Cancelled;
                    }

                }

                // Отправляем сообщение
                Task<bool> bitrSendMsgTask = Task<string[]>.Run(() =>
                {
                    return BitrixMessageSender.SendMsg_ToUsersChat(DBMainService.CurrentDBUser, selectedUsers.Select(se => se.DBUser), msg, imgId); ;
                });
                bitrSendMsgTask.Wait();


                if (bitrSendMsgTask.Result)
                {
                    MessageBox.Show($"Сообщение успешно отправлено! Открой Bitrix, там создан спец. чат", "KPLN", MessageBoxButtons.OK);
                    return Result.Succeeded;
                }

                return Result.Cancelled;
            }
            #endregion
#endif

            return Result.Cancelled;
        }
    }
}
