using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_TaskManager.Common;
using KPLN_TaskManager.ExecutableCommand;
using KPLN_TaskManager.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace KPLN_TaskManager.Forms
{
    public partial class TaskItemView : Window
    {
        public TaskItemView(TaskItemEntity taskItemEntity)
        {
            InitializeComponent();

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);

            CurrentTaskItemEntity = taskItemEntity;
            DataContext = CurrentTaskItemEntity;

            CurrentTaskComments = new ObservableCollection<TaskItemEntity_Comment>(TaskManagerDBService.GetComments_ByTaskItem(CurrentTaskItemEntity));
            TaskItem_Comments.ItemsSource = CurrentTaskComments;

            if (CurrentTaskItemEntity.BitrixTaskId == 0 || CurrentTaskItemEntity.BitrixTaskId == -1)
                BtnBitrixTask.Visibility = System.Windows.Visibility.Collapsed;

            
            if (CurrentTaskItemEntity.ElementIds != null && string.IsNullOrEmpty(CurrentTaskItemEntity.ElementIds) && CurrentTaskItemEntity.ModelName.Equals(Module.CurrentDocument))
            {
                ModelElemsContTBl.Visibility = System.Windows.Visibility.Collapsed;
                ModelViewIdTBl.Visibility = System.Windows.Visibility.Collapsed;
                SelectRevitElems.IsEnabled = false;
            }

            #region Настройка уровня доступа в зависимости от пользователя
            bool isNewTask = CurrentTaskItemEntity.Id == 0;
            bool isCreatorEditable = MainDBService.CurrentDBUserSubDepartment.Id == CurrentTaskItemEntity.CreatedTaskDepartmentId;
            bool isFullEditable = MainDBService.CurrentDBUserSubDepartment.Id == CurrentTaskItemEntity.CreatedTaskDepartmentId
                || MainDBService.CurrentDBUserSubDepartment.Id == CurrentTaskItemEntity.DelegatedDepartmentId;

            HeaderTBox.IsEnabled = isCreatorEditable;
            DelegateDepCBox.IsEnabled = isCreatorEditable;
            BodyTBox.IsReadOnly = !isCreatorEditable;
            ModelBindingCBox.IsEnabled = isCreatorEditable;
            ImgLoad.IsEnabled = isCreatorEditable;
            AddElementIdsBtn.IsEnabled = isCreatorEditable;
            RemoveElementIdsBtn.IsEnabled = isCreatorEditable;
            AddChangeTaskBtn.IsEnabled = isCreatorEditable;

            DelegateUserCBox.IsEnabled = isFullEditable;
            BitrixParentTaskIdTBox.IsEnabled = isFullEditable;

            BitrixExpander.IsEnabled = isFullEditable && !isNewTask;
            SendToBitrix.IsEnabled = isFullEditable && !isNewTask;
            ToggleStatusBtn.IsEnabled = isFullEditable && !isNewTask;
            CommentBtn.IsEnabled = isFullEditable && !isNewTask;
            #endregion
        }

        public TaskItemEntity CurrentTaskItemEntity { get; set; }

        public ObservableCollection<TaskItemEntity_Comment> CurrentTaskComments { get; set; }

        private static string GetCurrentData() => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        /// <summary>
        /// Событие наведения мыши в окно ScrollViewer. Связано с потеряй фокуса на колесо мыши
        /// </summary>
        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            CommentScroll.ScrollToVerticalOffset(CommentScroll.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void AddElementIdsBtn_Click(object sender, RoutedEventArgs e)
        {
            UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
            Document doc = uidoc.Document;

            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                MessageBox.Show(
                    "Чтобы добавить элементы к задаче - сначала выделите их в модели. ВАЖНО: Можно поставить задачу и добавить позже, путём изменения поставленной задачи",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Asterisk);

                return;
            }

            ModelElemsContTBl.Visibility = System.Windows.Visibility.Visible;
            ModelViewIdTBl.Visibility = System.Windows.Visibility.Visible;

            if (CurrentTaskItemEntity.Id != -1 && CurrentTaskItemEntity.Id != 0)
                SelectRevitElems.IsEnabled = true;

            CurrentTaskItemEntity.ElementIds = string.Join(",", selectedIds);
            CurrentTaskItemEntity.ModelViewId = uidoc.ActiveView.Id.IntegerValue;
        }

        private void RemoveElementIdsBtn_Click(object sender, RoutedEventArgs e)
        {
            ModelElemsContTBl.Visibility = System.Windows.Visibility.Collapsed;
            ModelViewIdTBl.Visibility = System.Windows.Visibility.Collapsed;
            SelectRevitElems.IsEnabled = false;

            CurrentTaskItemEntity.ModelName = string.Empty;
            CurrentTaskItemEntity.ElementIds = string.Empty;
        }

        /// <summary>
        /// Выделить элементы в Revit-модели 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectRevitElems_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentTaskItemEntity.ModelName))
            {
                MessageBox.Show(
                    "Элементы к задаче НЕ прикреплены",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Asterisk);

                return;
            }

            UIDocument uidoc = Module.CurrentUIApplication.ActiveUIDocument;
            Document doc = uidoc.Document;
            string currentDocName = TaskItemEntity.CurrentModelName(doc);

            if (CurrentTaskItemEntity.ModelName != currentDocName)
            {
                MessageBox.Show(
                    $"Элементы из задачи относятся к файлу \'{CurrentTaskItemEntity.ModelName}\', а у тебя открыт \'{currentDocName}\'",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return;
            }

            // Открываю вид. В Idling нельзя запихивать - не обработает: https://www.revitapidocs.com/2023/b6adb74b-39af-9213-c37b-f54db76b75a3.htm
            View viewFromTask = null;
            if (CurrentTaskItemEntity.ModelViewId != -1 || CurrentTaskItemEntity.ModelViewId != 0)
            {
                viewFromTask = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .WhereElementIsNotElementType()
                    .Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .FirstOrDefault(v => v.Id.IntegerValue == CurrentTaskItemEntity.ModelViewId);
            }
            if (viewFromTask != null)
                uidoc.ActiveView = viewFromTask;

            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandShowElement(CurrentTaskItemEntity));
            this.Close();
        }

        /// <summary>
        /// Создать/изменить задачу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddChangeTaskBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!VerifyMainFields())
                return;

            CurrentTaskItemEntity.LastChangeData = CurrentTaskItemEntity.CreatedTaskData;
            // Если id = 0 - это новая задача
            if (CurrentTaskItemEntity.Id == 0)
            {
                CurrentTaskItemEntity.CreatedTaskData = GetCurrentData();

                TaskManagerDBService.CreateDBTaskItem(CurrentTaskItemEntity);

                Module.MainMenuViewer.LoadTaskData();
            }
            // Иначе - задача уже была, нужно пересохранить
            else
            {
                TaskManagerDBService.UpdateTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);

                LogToCommentSender("Отредактировано");
            }

            this.Close();
        }

        private void CommentBtn_Click(object sender, RoutedEventArgs e)
        {
            UserTextInput userTextInput = new UserTextInput("Комментарий");
            bool? createNewComment = userTextInput.ShowDialog();
            if ((bool)createNewComment)
                LogToCommentSender(userTextInput.UserInput);
        }

        private void ToggleStatusBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button taskBtn = sender as System.Windows.Controls.Button;
            if (taskBtn.DataContext is TaskItemEntity tiEnt)
            {
                string inputFormTitle = string.Empty;
                string commetnStatus = string.Empty;
                switch (tiEnt.TaskStatus)
                {
                    case TaskStatusEnum.Open:
                        inputFormTitle = "Комментарий к закрытию";
                        commetnStatus = "Закрыто";
                        break;
                    case TaskStatusEnum.Close:
                        inputFormTitle = "Комментарий к возобновлению";
                        commetnStatus = "Возобновлено";
                        break;
                }

                UserTextInput userTextInput = new UserTextInput(inputFormTitle);
                bool? createNewComment = userTextInput.ShowDialog();
                if ((bool)createNewComment)
                {
                    switch (tiEnt.TaskStatus)
                    {
                        case TaskStatusEnum.Open:
                            tiEnt.TaskStatus = TaskStatusEnum.Close;
                            break;
                        case TaskStatusEnum.Close:
                            tiEnt.TaskStatus = TaskStatusEnum.Open;
                            break;
                    }
                    tiEnt.LastChangeData = GetCurrentData();
                    TaskManagerDBService.UpdateTaskItem_ByTaskItemEntity(tiEnt);

                    LogToCommentSender($"{commetnStatus}: {userTextInput.UserInput}");

                    this.Close();
                }
            }
        }

        /// <summary>
        /// Логирование действия в коммент к таске
        /// </summary>
        private void LogToCommentSender(string logMsg)
        {

            TaskItemEntity_Comment logComment = new TaskItemEntity_Comment(
                CurrentTaskItemEntity.Id,
                MainDBService.CurrentDBUser.Id,
                logMsg,
                GetCurrentData());

            TaskManagerDBService.CreateDBTaskItemComment(logComment);

            CurrentTaskComments.Insert(0, logComment);
        }

        /// <summary>
        /// Загрузка изображения из буфера
        /// </summary>
        private void ImgLoad_Click(object sender, RoutedEventArgs e)
        {
            if (Clipboard.ContainsImage())
            {
                BitmapSource image = Clipboard.GetImage();
                if (image != null)
                {
                    CurrentTaskItemEntity.ImageBuffer = ConvertToByteArray(image);

                    MessageBox.Show(
                        "Рисунок успешно добавлен!",
                        "Загрузка рисунка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);

                }
            }
            else
            {
                MessageBox.Show(
                    "Создай рисунок любым удобным способом, скопируй его в буфер обмена, а потом нажми \"Загрузить изображение\". ВАЖНО: Можно поставить задачу и добавить позже, путём изменения поставленной задачи",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BodyImg_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ImgLargeFrom imgLargeFrom = new ImgLargeFrom(CurrentTaskItemEntity);
            bool? imgLargeFromResult = imgLargeFrom.ShowDialog();
        }

        private void SendToBitrix_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            if (!VerifyMainFields())
            {
                this.Show();
                return;
            }

            if (CurrentTaskItemEntity.DelegatedTaskUserId == -1 || CurrentTaskItemEntity.DelegatedTaskUserId == 0)
            {
                MessageBox.Show(
                    $"Обязательно укажи исполнителя, прежде чем ставить задачу",
                    "Ошибка постановки задачи в Bitrix",
                    MessageBoxButton.OK,
                    MessageBoxImage.Asterisk);

                this.Show();
                return;
            }

            if (CurrentTaskItemEntity.BitrixParentTaskId == -1 || CurrentTaskItemEntity.BitrixParentTaskId == 0)
            {
                MessageBox.Show(
                    $"Обязательно укажи ID базовой задачи из Bitrix, прежде чем ставить новую задачу",
                    "Ошибка постановки задачи в Bitrix",
                    MessageBoxButton.OK,
                    MessageBoxImage.Asterisk);

                this.Show();
                return;
            }

            #region Получаю группу по родительской задаче
            Task<int> groupIdTask = Task<int>.Run(() =>
            {
                return BitrixMessageSender
                    .GetBitrixGroupId_ByTaskId(CurrentTaskItemEntity.BitrixParentTaskId);
            });

            int groupId = groupIdTask.Result;
            if (groupId == 0)
            {
                MessageBox.Show(
                    $"Не удалось найти проект для задачи с ID: {CurrentTaskItemEntity.BitrixParentTaskId}",
                    "Ошибка постановки задачи в Bitrix",
                    MessageBoxButton.OK,
                    MessageBoxImage.Asterisk);

                this.Show();
                return;
            }
            #endregion

            #region Создаю новую задачу
            Task<string> newTaskID = Task<string>.Run(() =>
            {
                string taskBodyWithElems;
                if (string.IsNullOrEmpty(CurrentTaskItemEntity.ModelName))
                    taskBodyWithElems = $"{CurrentTaskItemEntity.TaskBody}";
                else
                    taskBodyWithElems = $"{CurrentTaskItemEntity.TaskBody}\n\nИмя файла: {CurrentTaskItemEntity.ModelName}\nID элемента/-ов: {CurrentTaskItemEntity.ElementIds}";

                return BitrixMessageSender
                    .CreateTask_ByMainFields_AutoAuditors(
                        groupId,
                        CurrentTaskItemEntity.TaskTitle,
                        taskBodyWithElems,
                        CurrentTaskItemEntity.BitrixParentTaskId,
                        "BIM_Менеджер задач",
                        MainDBService.CurrentDBUser.BitrixUserID,
                        MainDBService.UserDbService.GetDBUser_ById(CurrentTaskItemEntity.DelegatedTaskUserId).BitrixUserID);
            });

            string newTaskIDResult = newTaskID.Result;

            if (newTaskIDResult == null
                || string.IsNullOrEmpty(newTaskIDResult)
                || !int.TryParse(newTaskID.Result, out int bitrTaskId))
            {
                MessageBox.Show(
                    $"Не удалось получить ID-задачи из Bitrix. Обратись к разработчику",
                    "Отправка в Bitrix",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Show();
                return;
            }

            MessageBox.Show(
                $"Задача успешно поставлена! ID задачи - {bitrTaskId}",
                "Отправка в Bitrix",
                MessageBoxButton.OK,
                MessageBoxImage.Asterisk);
            #endregion

            #region Фикисрую данные в БД
            CurrentTaskItemEntity.BitrixTaskId = bitrTaskId;
            BtnBitrixTask.Visibility = System.Windows.Visibility.Visible;
            TaskManagerDBService.UpdateTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);

            LogToCommentSender($"Создана задача с ID: {bitrTaskId}");
            #endregion

            #region Загрузка рисунка в задачу
            if (CurrentTaskItemEntity.ImageBuffer == null || CurrentTaskItemEntity.ImageBuffer.Length < 2)
            {
                this.Show();
                return;
            }

            string fileName = $"TaskManager_ImgForTask_{bitrTaskId}";
            Task<string> loadImgToBitrDisk = Task<string>.Run(() =>
            {
                return BitrixMessageSender.UploadFile(groupId, CurrentTaskItemEntity.ImageBuffer, fileName);
            });

            if (loadImgToBitrDisk.Result != null
                && !string.IsNullOrEmpty(loadImgToBitrDisk.Result)
                && int.TryParse(loadImgToBitrDisk.Result, out int imgIdFromBirt))
            {
                Task<bool> loadImgToBitixTask = Task<bool>.Run(() =>
                {
                    return BitrixMessageSender.UpdateTask_LoadImg(bitrTaskId, imgIdFromBirt);
                });

                if (loadImgToBitixTask.Result)
                    LogToCommentSender($"Рисунок к задаче с ID: {bitrTaskId} - успешно добавлен!");
                else
                {
                    MessageBox.Show(
                        $"Рисунок не удалось прикрепить к задаче! Отправь разработчику - не удалось прикрепить файл к задаче Битрикс",
                        "Отправка в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show(
                    $"Рисунок не удалось прикрепить к задаче! Отправь разработчику - не удалось получить ID файла из диска Битрикс",
                    "Отправка в Bitrix",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            #endregion

            this.Show();
        }

        private void BtnBitrixTask_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button taskBtn = sender as System.Windows.Controls.Button;
            if (taskBtn.DataContext is TaskItemEntity tiEnt)
            {
                if (MainDBService.CurrentDBUser.BitrixUserID == -1 || MainDBService.CurrentDBUser.BitrixUserID == 0)
                {
                    MessageBox.Show(
                        $"Не удалось получить ID-пользователя из Bitrix. Обратись к разработчику",
                        "Отправка в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                Process.Start("chrome", $"https://kpln.bitrix24.ru/company/personal/user/{MainDBService.CurrentDBUser.BitrixUserID}/tasks/task/view/{tiEnt.BitrixTaskId}/");
            }
        }

        private byte[] ConvertToByteArray(BitmapSource image)
        {
            byte[] resultBit;

            JpegBitmapEncoder encoder = new JpegBitmapEncoder();
            encoder.QualityLevel = 100;
            using (MemoryStream stream = new MemoryStream())
            {
                encoder.Frames.Add(BitmapFrame.Create(image));
                encoder.Save(stream);
                resultBit = stream.ToArray();
                stream.Close();
            }

            return resultBit;
        }

        /// <summary>
        /// Предварительная проверка данных основных полей
        /// </summary>
        /// <returns></returns>
        private bool VerifyMainFields()
        {
            if (string.IsNullOrEmpty(CurrentTaskItemEntity.TaskHeader)
                || string.IsNullOrEmpty(CurrentTaskItemEntity.TaskBody)
                || CurrentTaskItemEntity.DelegatedDepartmentId == -1
                || CurrentTaskItemEntity.DelegatedDepartmentId == 0)
            {
                MessageBox.Show(
                    $"Обязательно заполни заголовок, ответсвенный отдел и описание ошибки, прежде чем ставить задачу",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return false;
            }

            return true;
        }
    }
}
