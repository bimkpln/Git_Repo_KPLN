using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.Services;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker;
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
using System.Windows.Controls;
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

            CurrentTaskComments = new ObservableCollection<TaskItemEntity_Comment>(TMDBService.GetComments_ByTaskItem(CurrentTaskItemEntity));
            TaskItem_Comments.ItemsSource = CurrentTaskComments;

            // Видимость поля с задачей в битрикс
            if (CurrentTaskItemEntity.BitrixTaskId == 0 || CurrentTaskItemEntity.BitrixTaskId == -1)
                BtnBitrixTask.Visibility = System.Windows.Visibility.Collapsed;

            // Видимость и нажимаемость кнопки выбора эл-в
            if (CurrentTaskItemEntity.ElementIds != null && string.IsNullOrEmpty(CurrentTaskItemEntity.ElementIds) && CurrentTaskItemEntity.ModelName.Equals(Module.CurrentDoc))
            {
                ModelElemsContTBl.Visibility = System.Windows.Visibility.Collapsed;
                ModelViewIdTBl.Visibility = System.Windows.Visibility.Collapsed;
                SelectRevitElems.IsEnabled = false;
            }

            // Настройка экспандера рисунка
            SetImgExpander();

            // Настройка уровня доступа в зависимости от пользователя
            SetUserAccessLevel();

            // Подписка на обновление из БД при открытии
            this.Loaded += (s, e) => RefreshFromDatabase();
        }

        public TaskItemEntity CurrentTaskItemEntity { get; set; }

        public ObservableCollection<TaskItemEntity_Comment> CurrentTaskComments { get; set; }

        /// <summary>
        /// Таска редактируется создателем?
        /// </summary>
        public bool IsCreatorEditable { get; set; } = false;

        private static string GetCurrentData() => DateTime.Now.ToString("yyyy/MM/dd_HH:mm");

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        /// <summary>
        /// Обновление динамичных данных сущности на данные из БД
        /// </summary>
        private void RefreshFromDatabase()
        {
            TaskItemEntity freshTIE = TMDBService.GetEntity_ByEntityId(CurrentTaskItemEntity.Id);

            if (freshTIE == null)
                return;

            CurrentTaskItemEntity.TaskHeader = freshTIE.TaskHeader;
            CurrentTaskItemEntity.DelegatedDepartmentId = freshTIE.DelegatedDepartmentId;
            CurrentTaskItemEntity.BitrixTaskId = freshTIE.BitrixTaskId;
            CurrentTaskItemEntity.TaskBody = freshTIE.TaskBody;
            CurrentTaskItemEntity.PathToImageBufferDB = freshTIE.PathToImageBufferDB;
            CurrentTaskItemEntity.ModelName = freshTIE.ModelName;
            CurrentTaskItemEntity.ModelViewId = freshTIE.ModelViewId;
            CurrentTaskItemEntity.ElementIds = freshTIE.ElementIds;
            CurrentTaskItemEntity.TaskStatus = freshTIE.TaskStatus;
            CurrentTaskItemEntity.LastChangeData = freshTIE.LastChangeData;

            CurrentTaskComments?.Clear();
            foreach (var c in TMDBService.GetComments_ByTaskItem(CurrentTaskItemEntity))
                CurrentTaskComments.Add(c);
        }

        /// <summary>
        /// Настройка уровня доступа в зависимости от пользователя
        /// </summary>
        private void SetUserAccessLevel()
        {
            bool isDepSubDep = DBMainService.DBSubDepartmentColl.Any(sd => sd.DependentSubDepId == DBMainService.CurrentUserDBSubDepartment.Id);
            
            // Настройка выбора отдела ОТ в завиисмости от наличия подчиненных подотделов
            CreateDepCBox.IsEnabled = isDepSubDep;

            bool isNewTask = CurrentTaskItemEntity.Id == 0;
            IsCreatorEditable = DBMainService.CurrentUserDBSubDepartment.Id == CurrentTaskItemEntity.CreatedTaskDepartmentId || isDepSubDep;
            bool isFullEditable = DBMainService.CurrentUserDBSubDepartment.Id == CurrentTaskItemEntity.CreatedTaskDepartmentId
                || DBMainService.CurrentUserDBSubDepartment.Id == CurrentTaskItemEntity.DelegatedDepartmentId
                || isDepSubDep;

            HeaderTBox.IsEnabled = IsCreatorEditable;
            DelegateDepCBox.IsEnabled = IsCreatorEditable;
            BodyTBox.IsReadOnly = !IsCreatorEditable;
            ModelBindingCBox.IsEnabled = IsCreatorEditable;
            DeleteBtn.IsEnabled = IsCreatorEditable;
            ImgLoad.IsEnabled = IsCreatorEditable;
            AddElementIdsBtn.IsEnabled = IsCreatorEditable;
            RemoveElementIdsBtn.IsEnabled = IsCreatorEditable;
            AddChangeTaskBtn.IsEnabled = IsCreatorEditable;

            DelegateUserCBox.IsEnabled = isFullEditable;
            BitrixParentTaskIdTBox.IsEnabled = isFullEditable;

            BitrixExpander.IsEnabled = isFullEditable && !isNewTask;
            SendToBitrix.IsEnabled = (isFullEditable && !isNewTask) && CurrentTaskItemEntity.BitrixTaskId == -1;
            ToggleStatusBtn.IsEnabled = isFullEditable && !isNewTask;
            CommentBtn.IsEnabled = isFullEditable && !isNewTask;
        }

        /// <summary>
        /// Дополнительный триггер на замену отдела для управления уровнем доступа
        /// </summary>
        private void CreateDepCBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e) => SetUserAccessLevel();

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

            List<ElementId> selectedIds = new List<ElementId>();
            bool isElemFromDoc = true;
            Selection selection = uidoc.Selection;
#if Revit2020 || Debug2020
            // Для 2020 - нет возможности получить элементы из связи. Это исправили в более поздних версиях
            selectedIds = uidoc.Selection.GetElementIds().ToList();
#else
            IList<Reference> selRefers = selection.GetReferences();
            foreach (Reference selRefer in selRefers)
            {
                Element selElem = uidoc.Document.GetElement(selRefer);
                if (selElem is RevitLinkInstance rli)
                {
                    Document linkedDoc = rli.GetLinkDocument();
                    Element linkElem = linkedDoc.GetElement(selRefer.LinkedElementId);
                    if (linkElem != null)
                    {
                        selectedIds.Add(linkElem.Id);
                        isElemFromDoc = false;
                    }
                }
                else
                    selectedIds.Add(selElem.Id);
            }
#endif


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
            if (isElemFromDoc)
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
            if (string.IsNullOrEmpty(CurrentTaskItemEntity.ModelName) || CurrentTaskItemEntity.ElementIdsCount <= 0)
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
        }

        /// <summary>
        /// Создать/изменить задачу
        /// </summary>
        private void AddChangeTaskBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!VerifyMainFields())
                return;

            // Если id = 0 - это новая задача
            if (CurrentTaskItemEntity.Id == 0)
            {
                int nextIdFromDB = TMDBService.GetNextId();
                
                CurrentTaskItemEntity.CreatedTaskData = GetCurrentData();
                CurrentTaskItemEntity.LastChangeData = GetCurrentData();
                CurrentTaskItemEntity.Id = nextIdFromDB;
                CurrentTaskItemEntity.PathToImageBufferDB = TM_IBDBService.GenerateNewPath_DBForTMIB(CurrentTaskItemEntity).FullName;

                Task createItemTask = Task.Run(() =>
                {
                    TMDBService.CreateDBTaskItem(CurrentTaskItemEntity);
                    TM_IBDBService.CreateDbFile_ByTaskItem(CurrentTaskItemEntity);
                });
                createItemTask.Wait();

                if(!CurrentTaskItemEntity.TE_ImageBufferColl.All(buff => buff.ImageBuffer == null || buff.ImageBuffer.Length == 0))
                    TM_IBDBService.CreateDBTaskEntity_ImageBufferItem(CurrentTaskItemEntity);

                Module.MainMenuViewer.LoadTaskData();
            }
            // Иначе - задача уже была, нужно пересохранить
            else
            {
                CurrentTaskItemEntity.LastChangeData = GetCurrentData();

                // Создаю БД, если ранее не было (связано со старыми замечаниями, где картинки хранились в одной БД). !!!УДАЛИТЬ 01.10.2026!!!
                if (string.IsNullOrEmpty(CurrentTaskItemEntity.PathToImageBufferDB))
                {
                    CurrentTaskItemEntity.PathToImageBufferDB = TM_IBDBService.GenerateNewPath_DBForTMIB(CurrentTaskItemEntity).FullName;
                    TM_IBDBService.CreateDbFile_ByTaskItem(CurrentTaskItemEntity);
                }

                Task updateItemTask = Task.Run(() =>
                {
                    TMDBService.UpdateTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);
                    TM_IBDBService.UpSetTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);
                });
                updateItemTask.Wait();

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
            if (CurrentTaskItemEntity != null)
            {
                string inputFormTitle = string.Empty;
                string commetnStatus = string.Empty;
                switch (CurrentTaskItemEntity.TaskStatus)
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
                    switch (CurrentTaskItemEntity.TaskStatus)
                    {
                        case TaskStatusEnum.Open:
                            CurrentTaskItemEntity.TaskStatus = TaskStatusEnum.Close;
                            break;
                        case TaskStatusEnum.Close:
                            CurrentTaskItemEntity.TaskStatus = TaskStatusEnum.Open;
                            break;
                    }

                    CurrentTaskItemEntity.LastChangeData = GetCurrentData();
                    TMDBService.UpdateTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);

                    LogToCommentSender($"{commetnStatus}: {userTextInput.UserInput}");
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
                DBMainService.CurrentDBUser.Id,
                logMsg,
                GetCurrentData());

            TMDBService.CreateDBTaskItemComment(logComment);

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
                    SetImageBuffer(image);
                    SetImgExpander();

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

        public void ImgLeft_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentTaskItemEntity.TE_ImageBuffer_Current == 0)
                CurrentTaskItemEntity.TE_ImageBuffer_Current = CurrentTaskItemEntity.TE_ImageBufferColl.Count - 1;
            else
                CurrentTaskItemEntity.TE_ImageBuffer_Current--;

            Button btn = (Button)sender;
            if (btn.DataContext is ImgLargeFrom imgForm)
                imgForm.ILF_TaskImageSource = CurrentTaskItemEntity.TaskImageSource;
        }

        public void ImgRight_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentTaskItemEntity.TE_ImageBuffer_Current == CurrentTaskItemEntity.TE_ImageBufferColl.Count - 1)
                CurrentTaskItemEntity.TE_ImageBuffer_Current = 0;
            else
                CurrentTaskItemEntity.TE_ImageBuffer_Current++;
            
            
            Button btn = (Button)sender;
            if (btn.DataContext is ImgLargeFrom imgForm)
                imgForm.ILF_TaskImageSource = CurrentTaskItemEntity.TaskImageSource;
        }

        public void ImgDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult astericsResult = MessageBox.Show(
                $"Удалить текущий рисунок?",
                "Внимание",
                MessageBoxButton.YesNo,
                MessageBoxImage.Asterisk);

            if (astericsResult == MessageBoxResult.No)
                return;

            DeleteCurrentImageBuffer();

            Button btn = (Button)sender;
            if (btn.DataContext is ImgLargeFrom imgForm)
            {
                imgForm.ILF_TaskImageSource = CurrentTaskItemEntity.TaskImageSource;
                imgForm.ILF_TE_ImageBufferColl = CurrentTaskItemEntity.TE_ImageBufferColl;
            }
        }

        private void BodyImg_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            ImgLargeFrom imgLargeFrom = new ImgLargeFrom(this, CurrentTaskItemEntity);
            WindowHandleSearch.MainWindowHandle.SetAsOwner(imgLargeFrom);

            this.Hide();
            imgLargeFrom.ShowDialog();
            this.Show();
        }

        [Obsolete("!!!НУЖНО ПРОТЕСТИРОВАТЬ!!!")]
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
                        DBMainService.CurrentDBUser.BitrixUserID,
                        DBMainService.UserDbService.GetDBUser_ById(CurrentTaskItemEntity.DelegatedTaskUserId).BitrixUserID);
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
            CurrentTaskItemEntity.LastChangeData = GetCurrentData();
            BtnBitrixTask.Visibility = System.Windows.Visibility.Visible;
            TMDBService.UpdateTaskItem_ByTaskItemEntity(CurrentTaskItemEntity);

            LogToCommentSender($"Создана задача с ID: {bitrTaskId}");
            #endregion

            #region Загрузка рисунка в задачу
            if (CurrentTaskItemEntity.TE_ImageBufferColl == null 
                || CurrentTaskItemEntity.TE_ImageBufferColl.All(buff => buff.ImageBuffer == null || buff.ImageBuffer.Length == 0))
            {
                this.Show();
                return;
            }


            // Гружу картинки в битрикс
            Task<string>[] bitrUploadTasks = new Task<string>[CurrentTaskItemEntity.TE_ImageBufferColl.Count];
            for (int i = 0; i < CurrentTaskItemEntity.TE_ImageBufferColl.Count; i++)
            {
                string fileName = $"TaskManager_ImgForTask_{bitrTaskId}_{i}";

                byte[] imgBuff = CurrentTaskItemEntity.TE_ImageBufferColl[i].ImageBuffer;
                if (imgBuff.Length < 2)
                    continue;

                bitrUploadTasks[i] = Task<string>.Run(() =>
                {
                    return BitrixMessageSender.UploadFile(groupId, imgBuff, fileName);
                });

            }
            Task.WaitAll(bitrUploadTasks);


            // Привязывю картинки с задачей
            foreach(Task<string> bitrUploadTask in bitrUploadTasks)
            {
                string loadImgToBitrDisk = bitrUploadTask.Result;
                if (loadImgToBitrDisk != null
                    && !string.IsNullOrEmpty(loadImgToBitrDisk)
                    && int.TryParse(loadImgToBitrDisk, out int imgIdFromBirt))
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
            }
            #endregion

            this.Show();
        }

        private void BtnBitrixTask_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button taskBtn = sender as System.Windows.Controls.Button;
            if (taskBtn.DataContext is TaskItemEntity tiEnt)
            {
                if (DBMainService.CurrentDBUser.BitrixUserID == -1 || DBMainService.CurrentDBUser.BitrixUserID == 0)
                {
                    MessageBox.Show(
                        $"Не удалось получить ID-пользователя из Bitrix. Обратись к разработчику",
                        "Отправка в Bitrix",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return;
                }

                Process.Start("chrome", $"https://kpln.bitrix24.ru/company/personal/user/{DBMainService.CurrentDBUser.BitrixUserID}/tasks/task/view/{tiEnt.BitrixTaskId}/");
            }
        }

        /// <summary>
        /// Установить ImageBuffer
        /// </summary>
        private void SetImageBuffer(BitmapSource image)
        {
            byte[] resultBit;
            JpegBitmapEncoder encoder = new JpegBitmapEncoder
            {
                QualityLevel = 100
            };

            using (MemoryStream stream = new MemoryStream())
            {
                encoder.Frames.Add(BitmapFrame.Create(image));
                encoder.Save(stream);
                resultBit = stream.ToArray();
                stream.Close();
            }

            TaskEntity_ImageBuffer newTE_IMB = new TaskEntity_ImageBuffer(CurrentTaskItemEntity.TE_ImageBufferColl.Count, CurrentTaskItemEntity.Id, resultBit);
            CurrentTaskItemEntity.TE_ImageBufferColl.Add(newTE_IMB);
            CurrentTaskItemEntity.TE_ImageBuffer_Current = CurrentTaskItemEntity.TE_ImageBufferColl.Count - 1;
        }

        /// <summary>
        /// Удалить текущий ImageBuffer
        /// </summary>
        private void DeleteCurrentImageBuffer()
        {
            TaskEntity_ImageBuffer currentIB = CurrentTaskItemEntity.TE_ImageBufferColl[CurrentTaskItemEntity.TE_ImageBuffer_Current];
            CurrentTaskItemEntity.TE_ImageBufferColl.Remove(currentIB);
            CurrentTaskItemEntity.TE_ImageBuffer_Current = CurrentTaskItemEntity.TE_ImageBufferColl.Count - 1;
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
                    $"Обязательно заполни заголовок, ответственный отдел и описание ошибки, прежде чем ставить задачу",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return false;
            }

            return true;
        }

        /// <summary>
        /// Настройка заголовка и "свёрнутости" Expander с изображением
        /// </summary>
        private void SetImgExpander()
        {
            if (CurrentTaskItemEntity.TaskImageSource != null)
            {
                ImgExpander.IsEnabled = true;
                ImgExpander.IsExpanded = true;
            }
            else
            {
                ImgExpander.IsEnabled= false;
                ImgExpander.IsExpanded = false;
            }
        }        
    }
}
