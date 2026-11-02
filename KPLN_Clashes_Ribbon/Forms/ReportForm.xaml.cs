using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Services.SQLite;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.ExternalEventHandler;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_DBWorker;
using KPLN_Library_Forms.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportItemFrom.xaml
    /// </summary>
    public partial class ReportForm : Window
    {
        private static ReportForm _currentInstance;
        private static bool _isOpening;

        public List<IExecutableCommand> OnClosingActions = new List<IExecutableCommand>();

        private readonly ReportManagerForm _reportManagerForm;
        private readonly ReportGroup _repourtGroup;
        private readonly Report _currentReport;
        private readonly Services.SQLite.SQLiteService_ReportItemsDB _sqliteService_ReportInstanceDB;
        private readonly Services.SQLite.SQLiteService_MainDB _sqliteService_MainDB = new Services.SQLite.SQLiteService_MainDB();

        private readonly ViewActivatedHandler _viewHandler;

        private bool _isSaving;

        private string _conflictDataTBx = string.Empty;
        private string _idDataTBx = string.Empty;
        private string _conflictMetaDataTBx = string.Empty;

        public static void ShowOrActivate(ReportManagerForm reportManagerForm, Report report, ReportGroup reportGroup, ViewActivatedHandler viewHandler)
        {
            // Общее окно для всех менеджеров и документов в этом процессе Revit.
            if (_currentInstance != null)
            {
                if (_currentInstance.WindowState == WindowState.Minimized)
                    _currentInstance.WindowState = WindowState.Normal;
                if (!_currentInstance.IsVisible)
                    _currentInstance.Show();

                _currentInstance.Activate();
                _currentInstance.Focus();
                return;
            }

            // Защита от повторного открытия, пока конструктор загружает отчет.
            if (_isOpening)
                return;

            _isOpening = true;
            try
            {
                _currentInstance = new ReportForm(reportManagerForm, report, reportGroup, viewHandler);
                _currentInstance.Show();
            }
            catch
            {
                _currentInstance = null;
                throw;
            }
            finally
            {
                _isOpening = false;
            }
        }

        private ReportForm(ReportManagerForm reportManagerForm, Report report, ReportGroup reportGroup, ViewActivatedHandler viewHandler)
        {
            _reportManagerForm = reportManagerForm;
            _repourtGroup = reportGroup;
            _currentReport = report;
            _viewHandler = viewHandler;


            _sqliteService_ReportInstanceDB = new Services.SQLite.SQLiteService_ReportItemsDB(_currentReport.PathToReportInstance);

            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();
            FilteredInstancesColl = CollectionViewSource.GetDefaultView(ReportInstancesColl);
            FilteredInstancesColl.Filter += FilterForRepItems;
            
            
            

            UIApplication uiapp = _viewHandler.CurrnetUIApplication;
            Document doc = uiapp?.ActiveUIDocument?.Document;


            if (_repourtGroup.IsEnabled)
            {
                foreach (ReportItem instance in ReportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Visible;
                    instance.IsControllsEnabled = true;
                }
            }
            else
            {
                foreach (ReportItem instance in ReportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Collapsed;
                    instance.IsControllsEnabled = false;
                }
            }

            InitializeComponent();


            this.Title = string.Format("KPLN: Отчет Navisworks ({0})", report.Name);
            this.Owner = reportManagerForm;
            this.AllItemsRB.IsChecked = true;
            this.DataContext = this;


            OnClosingActions.Add(new CmdRemoveFamInst());
            Closing += RemoveOnClose;
            StartLiveUpdates();
            Closed += (sender, args) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        private bool FilterForRepItems(object obj)
        {
            if (obj is ReportItem item)
            {
                System.Windows.Controls.ComboBox sfCBX = this.StatusFilterCBX;
                if (sfCBX == null)
                    return true;

                // Ключевя переменная по имени файла
                bool docCheck = false;
                if ((bool)this.AllItemsRB.IsChecked)
                    docCheck = true;
                else
                {
                    UIApplication uiapp = _viewHandler.CurrnetUIApplication;
                    if (uiapp == null || uiapp.ActiveUIDocument == null)
                        docCheck = false;
                    else
                    {
                        string fileFullName = KPLN_Library_DBWorker.FactoryParts.SQLite.SQLiteDocService.GetFileFullName(uiapp.ActiveUIDocument.Document);
                        string docTitle = fileFullName.Split('/').LastOrDefault().Split(new[] {".rvt"}, StringSplitOptions.None).FirstOrDefault();
                        if (string.IsNullOrEmpty(docTitle))
                            return true;
                        
                        foreach(ReportItem subItem in item.SubElements)
                        {
                            if (docTitle.Contains(subItem.Element_1_DocName) || docTitle.Contains(subItem.Element_2_DocName))
                            {
                                docCheck = true;
                                break;
                            }
                        }
                    }
                }


                // Ключевя переменная по статусу
                bool statusCheck = false;
                int mainStatusFilterIndex = this.StatusFilterCBX.SelectedIndex;
                switch (mainStatusFilterIndex)
                {
                    case 0:
                        statusCheck = true;
                        break;
                    case 1:
                        if (item.Status == KPItemStatus.Opened || item.Status == KPItemStatus.Delegated)
                            statusCheck = true;

                        break;
                    case 2:
                        if (item.Status == KPItemStatus.Opened)
                            statusCheck = true;

                        break;
                    case 3:
                        if (item.Status == KPItemStatus.Delegated)
                            statusCheck = true;

                        break;
                    case 4:
                        if (item.Status == KPItemStatus.Approved)
                            statusCheck = true;

                        break;
                    case 5:
                        if (item.Status == KPItemStatus.Closed)
                            statusCheck = true;

                        break;
                }

                return docCheck && statusCheck && CheckParamData(item);
            }

            return false;
        }

        /// <summary>
        /// Фильтрация элемента по атрибутам
        /// </summary>
        private bool CheckParamData(ReportItem item)
        {
            // Фильтрация по атрибутам
            bool isEmptyConflData = string.IsNullOrEmpty(_conflictDataTBx);
            bool isEmptyConflictMetaData = string.IsNullOrEmpty(_conflictMetaDataTBx);
            bool isEmptyIDData = string.IsNullOrEmpty(_idDataTBx);

            if (isEmptyConflData && isEmptyConflictMetaData && isEmptyIDData)
                return true;
            else
            {
                bool checkConflData = item.Name.IndexOf(_conflictDataTBx, StringComparison.OrdinalIgnoreCase) >= 0
                    || item.SubElements.Any(sub => sub.Name.IndexOf(_conflictDataTBx, StringComparison.OrdinalIgnoreCase) >= 0);

                bool checkConflictMetaData = item.SubElements.Any(sub => sub.Element_1_Info?.IndexOf(_conflictMetaDataTBx, StringComparison.OrdinalIgnoreCase) >= 0
                        || sub.Element_2_Info?.IndexOf(_conflictMetaDataTBx, StringComparison.OrdinalIgnoreCase) >= 0);

                bool checkIDData = isEmptyIDData || item.GroupElementIds.Contains(_idDataTBx);

                if (checkConflData && checkConflictMetaData && checkIDData)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Исходная коллекция элементов
        /// </summary>
        public ReportItem[] ReportInstancesColl { get; private set; }

        /// <summary>
        /// Отфильтрованная коллекция элементов, которая является контекстом для окна
        /// </summary>
        public ICollectionView FilteredInstancesColl { get; private set; }

        private void RemoveOnClose(object sender, CancelEventArgs args)
        {
            if (_isSaving)
            {
                args.Cancel = true;
                return;
            }

            foreach (IExecutableCommand cmd in OnClosingActions)
            {
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(cmd);
            }

            _reportManagerForm.UpdateSelectedReportGroup(_repourtGroup);
        }

        private void OnLoadImage(object sender, RoutedEventArgs e)
        {
            if ((sender as System.Windows.Controls.Button).DataContext is ReportItem reportItem)
            {
                byte[] image_bytes = _sqliteService_ReportInstanceDB.GetImageBytes_ByItem(reportItem);
                if (image_bytes != null)
                    reportItem.LoadImage(image_bytes);
            }
        }

        /// <summary>
        /// Суммирование всех вызываемых элементов по кнопке
        /// </summary>
        private void SelectIdElements(object sender, RoutedEventArgs e)
        {
            ReportItem report = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            SelectIds(sender, report);
        }

        /// <summary>
        /// Разделение вызываемых элементов по кнопке для конфликта №1
        /// </summary>
        private void SelectIdElement_1(object sender, RoutedEventArgs e)
        {
            ReportItem report = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            SelectId(sender, report.Element_1_Info);
        }

        /// <summary>
        /// Разделение вызываемых элементов по кнопке для конфликта №2
        /// </summary>
        private void SelectIdElement_2(object sender, RoutedEventArgs e)
        {
            ReportItem report = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            SelectId(sender, report.Element_2_Info);
        }

        /// <summary>
        /// Выбор по списку id
        /// </summary>
        private void SelectIds(object sender, ReportItem report)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            var ids = report.GroupElementIds
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(strId => int.Parse(strId));
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CmdZoomSelectElems<int>(ids));
#else
            var ids = report.GroupElementIds
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(strId => long.Parse(strId));
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CmdZoomSelectElems<long>(ids));
#endif
        }

        /// <summary>
        /// Выбор по одиночному id
        /// </summary>
        private void SelectId(object sender, string elInfo)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            if (int.TryParse((sender as System.Windows.Controls.Button).Content.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CmdZoomSelectElems<int>(id, elInfo));
#else
            if (long.TryParse((sender as System.Windows.Controls.Button).Content.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long id))
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CmdZoomSelectElems<long>(id, elInfo));
#endif
            else
                throw new Exception("Проблемы с отчетом: параметр id не парсится");
        }

        private void PlacePoint(object sender, RoutedEventArgs args)
        {
            if ((sender as System.Windows.Controls.Button).DataContext is ReportItem report)
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CmdPlaceFamInst(report));
        }

        private void StatusFilterCBX_SelectionChanged(object sender, SelectionChangedEventArgs e) => FilteredInstancesColl?.Refresh();

        private void ConflictDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _conflictDataTBx = textBox.Text;

            FilteredInstancesColl?.Refresh();
        }

        private void IDDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _idDataTBx = textBox.Text;

            FilteredInstancesColl?.Refresh();
        }

        private void ConflictMetaDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _conflictMetaDataTBx = textBox.Text;

            FilteredInstancesColl?.Refresh();
        }

        private async void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            await PollChangesAsync(true);
        }
        private void OnZoomSettings(object sender, RoutedEventArgs e)
        {
            if (ZoomSettingsForm.TryActivateExisting())
                return;

            Document doc = _viewHandler?.CurrnetUIApplication?.ActiveUIDocument?.Document;
            if (doc == null)
            {
                System.Windows.MessageBox.Show(this, "Не удалось получить активный документ Revit для shared-конфига.", "Настройки зума", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ZoomSettingsForm form = new ZoomSettingsForm(doc) { Owner = this };
            form.ShowDialog();
        }

        /// <summary>Один пакет записи в фоне; UI меняется только после commit.</summary>
        private async Task<bool> SaveItemsAsync(List<ReportItem> items, KPItemStatus? status, string message, int? departmentId = null)
        {
            if (_isSaving || items.Count == 0)
                return false;

            var changes = items.Select(item => new SQLiteService_ReportItemsDB.ItemChange
            {
                Id = item.Id,
                DepartmentId = departmentId ?? item.DelegatedDepartmentId,
                Comment = new ReportItemComment(message).ToString(),
                Actor = SQLiteMainService.CurrentDBUser.Name + " " + SQLiteMainService.CurrentDBUser.Surname,
                SessionId = _reportSessionId
            }).ToArray();
            int reportId = _currentReport.Id;
            int groupId = _currentReport.ReportGroupId;
            string userName = SQLiteMainService.CurrentDBUser.SystemName;
            string fullName = SQLiteMainService.CurrentDBUser.Name + " " + SQLiteMainService.CurrentDBUser.Surname;
            string originalTitle = Title;
            var content = Content as UIElement;
            bool wasEnabled = content?.IsEnabled ?? true;
            _isSaving = true;
            _localSaveGeneration++;
            if (content != null)
                content.IsEnabled = false;
            Title = originalTitle + " — Сохранение…";
            bool itemsSaved = false;
            SQLiteService_ReportItemsDB.SavedItem[] saved = null;

            try
            {
                saved = await _sqliteService_ReportInstanceDB.SaveChangesBatchAsync(changes, status);
                itemsSaved = true;
                ApplyLiveRows(saved);

                return true;
            }
            catch (Exception ex)
            {
                string messageText = itemsSaved
                    ? "Изменения сохранены, но не удалось обновить окно. Откройте отчет заново; повторять запись не нужно.\n"
                    : "Не удалось сохранить изменения.\n";
                System.Windows.MessageBox.Show(this, messageText + ex.Message,
                    "Сохранение отчета", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            finally
            {
                _isSaving = false;
                Title = originalTitle;
                if (content != null)
                    content.IsEnabled = wasEnabled;

                if (saved != null)
                {
                    QueueMetadataRefresh(saved, status.HasValue, reportId, groupId, userName, fullName);
                    QueueCommentNames(saved);
                }
            }
        }

        /// <summary>
        /// Метод получения выделенных элементов
        /// </summary>
        private List<ReportItem> GetTargetItems(ReportItem clicked)
        {
            List<ReportItem> items = new List<ReportItem>();
            if (ReportControll.SelectedItems != null && ReportControll.SelectedItems.Count > 1)
            {
                foreach (object selected in ReportControll.SelectedItems)
                {
                    if (selected is ReportItem ri)
                        items.Add(ri);
                }
            }

            if (!items.Any(i => i.Id == clicked.Id))
                items.Add(clicked);

            if (items.Count > 1)
            {
                TaskDialog td = new TaskDialog("Внимание!")
                {
                    MainContent = $"Вы сейчас произведёте дейтсвия сразу с {items.Count} элементами. Продолжить?",
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning
                };

                TaskDialogResult tdResult;
                _isEditing = true;
                try
                {
                    tdResult = td.Show();
                }
                finally
                {
                    _isEditing = false;
                }
                if (tdResult == TaskDialogResult.No)
                    return new List<ReportItem>();
            }

            return items;
        }

        private async void OnCorrected(object sender, RoutedEventArgs e)
        {
            if (_isSaving)
                return;
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            await SaveItemsAsync(GetTargetItems(item), KPItemStatus.Closed, "Статус изменен: <Устранено>\n");
        }

        private async void OnApproved(object sender, RoutedEventArgs e)
        {
            if (_isSaving)
                return;
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            UserTextInput textInputForm = new UserTextInput("Введите комментарий для допуска:");
            if (ShowCommentInput(textInputForm))
                await SaveItemsAsync(GetTargetItems(item), KPItemStatus.Approved,
                    "Статус изменен: <Допустимое>\n" + textInputForm.UserInput);
        }

        private async void OnReset(object sender, RoutedEventArgs e)
        {
            if (_isSaving)
                return;
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            var items = GetTargetItems(item).Where(ri => ri.Status == KPItemStatus.Closed || ri.Status == KPItemStatus.Approved).ToList();
            await SaveItemsAsync(items, KPItemStatus.Opened, "Статус изменен: <Возвращен в работу - отказ в допуске>\n");
        }

        private async void OnAddComment(object sender, RoutedEventArgs e)
        {
            if (_isSaving)
                return;
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            UserTextInput textInputForm = new UserTextInput("Введите комментарий:");
            if (ShowCommentInput(textInputForm))
                await SaveItemsAsync(GetTargetItems(item), null, textInputForm.UserInput);
        }

        /// <summary>
        /// Кнопка делегирования коллизии на другой отдел
        /// </summary>
        private async void OnDelegate(object sender, RoutedEventArgs e)
        {
            if (_isSaving)
                return;
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            SubDepartmentBtn subDepartmentBtn = button.DataContext as SubDepartmentBtn;
            if (!(FindParent<ItemsControl>(button).DataContext is ReportItem currentItem))
                return;

            List<ReportItem> items = GetTargetItems(currentItem);

            if (items.Count == 0)
                return;

            if (subDepartmentBtn.Id == 99)
            {
                await SaveItemsAsync(items, KPItemStatus.Opened,
                    "Статус изменен: <Возвращен в работу - отказ в делегировании>\n", -1);
                return;
            }

            if (!await SaveItemsAsync(items, KPItemStatus.Delegated,
                $"Статус изменен: <Делегирована отделу {subDepartmentBtn.Name}>\n", subDepartmentBtn.Id))
                return;

            // Выделение и логирования делегирования при нажатии на кнопку сброса (по id)
            int delegBitrixTaskId = 0;
            switch (subDepartmentBtn.Id)
            {
                case 2:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdAR;
                    break;
                case 3:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdKR;
                    break;
                case 4:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdOV;
                    break;
                case 5:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdVK;
                    break;
                case 6:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdEOM;
                    break;
                case 7:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdSS;
                    break;
                case 20:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdITP;
                    break;
                case 21:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdAUPT;
                    break;
                case 22:
                    delegBitrixTaskId = _repourtGroup.BitrixTaskIdAV;
                    break;
            }

            // Сеть не блокирует поток интерфейса; сохранение замечаний уже подтверждено.
            try
            {
                if (delegBitrixTaskId > 0 && !await BitrixMessageSender.CheckTaskOpens_ByTaskId(delegBitrixTaskId))
                {
                    bool sent = await BitrixMessageSender.SendMsgToTask_ByTaskId(delegBitrixTaskId,
                        $"Пользователь <{SQLiteMainService.CurrentDBUser.Name} {SQLiteMainService.CurrentDBUser.Surname}> " +
                        $"делегировал вам коллизию из отчета: \"{_currentReport.Name}\" " +
                        $"с именем: \"{currentItem.Name}\"");
                    if (sent)
                        System.Windows.MessageBox.Show(
                            $"Отдел уже отработал свою задачу. Сообщение о делегировании отправлено в Bitrix: {delegBitrixTaskId}.",
                            "Bitrix", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Делегирование сохранено, но проверка/уведомление Bitrix не выполнено.\n" + ex.Message,
                    "Bitrix", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Сброс цвета кнопок делегирования
        /// </summary>
        private void ResetDelegateBtnBrush(ReportItem report)
        {
            ObservableCollection<SubDepartmentBtn> subDepartmentBtns = report.SubDepartmentBtns;
            foreach (SubDepartmentBtn sdBtn in subDepartmentBtns)
            {
                sdBtn.DelegateBtnBackground = Brushes.Transparent;
            }

            report.DelegatedDepartmentId = -1;
        }

        /// <summary>
        /// Переключение цветов по нажатию на кнопку делегирования
        /// </summary>
        private void SetDelegateBtnBrush(ReportItem report, SubDepartmentBtn btn)
        {
            ObservableCollection<SubDepartmentBtn> subDepartmentBtns = report.SubDepartmentBtns;
            foreach (SubDepartmentBtn sdBtn in subDepartmentBtns)
            {
                if (sdBtn.Id == btn.Id)
                {
                    btn.DelegateBtnBackground = Brushes.Aqua;
                    report.DelegatedDepartmentId = sdBtn.Id;
                }
                else
                    sdBtn.DelegateBtnBackground = Brushes.Transparent;
            }
        }

        private void OnExport(object sender, RoutedEventArgs e)
        {
            List<string> rows = new List<string>();
            foreach (object obj in FilteredInstancesColl)
            {
                if (obj is ReportItem instance)
                {
                    rows.Add(string.Empty);
                    rows.Add(GetExportRow(instance));
                    foreach (ReportItem subInstance in instance.SubElements)
                    {
                        rows.Add(GetExportRow(subInstance, instance));
                    }
                }
            }

            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Сохранить отчет",
                Filter = "txt file (*.txt)|*.txt",
                RestoreDirectory = true,
                FileName = string.Format("{0}_{1}-{2}-{3}.txt", _currentReport.Name.Replace('.', '_'), DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString()),
                CreatePrompt = false,
                OverwritePrompt = true,
                SupportMultiDottedExtensions = false
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(dialog.OpenFile());
                foreach (string row in rows)
                {
                    writer.WriteLine(row);
                }

                writer.Dispose();
                writer.Close();
            }
        }

        private static string Optimize(string value)
        {
            string final = string.Empty;
            string result = value.Replace("\n", "");
            result = result.Replace("\r", "");
            result = result.Replace("\t", "");
            bool lastCharIsSpace = false;
            foreach (char c in result)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (lastCharIsSpace)
                    {
                        lastCharIsSpace = true;
                    }
                    else
                    {
                        final += c;
                        lastCharIsSpace = true;
                    }
                }
                else if (char.IsSeparator(c))
                {
                    continue;
                }
                else
                {
                    final += c;
                    lastCharIsSpace = false;
                }
            }
            return final;
        }

        private static string GetExportRow(ReportItem instance, ReportItem parent = null)
        {

            List<string> parts = new List<string>
            {
                Optimize(instance.Id.ToString()),
                Optimize(instance.ParentGroupId.ToString()),
                Optimize(instance.Name)
            };
            if (parent != null)
            {
                parts.Add(Optimize(parent.Status.ToString("G")));
                if (parent.CommentCollection.Count != 0)
                {
                    ReportItemComment comment = parent.CommentCollection.Last();
                    parts.Add(string.Format("<{0}> {1} {2}", Optimize(comment.Time), Optimize(comment.UserFullName), Optimize(comment.Message)));
                }
                else
                {
                    parts.Add(Optimize("..."));
                }
            }
            else
            {
                parts.Add(Optimize(instance.Status.ToString("G")));
                if (instance.CommentCollection.Count != 0)
                {
                    ReportItemComment comment = instance.CommentCollection.Last();
                    parts.Add(string.Format("<{0}> {1} {2}", Optimize(comment.Time), Optimize(comment.UserFullName), Optimize(comment.Message)));
                }
                else
                {
                    parts.Add(Optimize("..."));
                }
            }

            if (instance.Element_1_Id != -1 && instance.Element_2_Id != -1)
            {
                parts.Add(Optimize(instance.Element_1_Id.ToString()));
                parts.Add(Optimize(instance.Element_1_Info));
                parts.Add(Optimize(instance.Element_2_Id.ToString()));
                parts.Add(Optimize(instance.Element_2_Info));
                parts.Add(Optimize(instance.Point));
            }

            return string.Join("\t", parts);
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            while (parentObject != null && !(parentObject is T))
            {
                parentObject = VisualTreeHelper.GetParent(parentObject);
            }

            return parentObject as T;
        }

        private void FileNameFilterGroup_Checked(object sender, RoutedEventArgs e) => FilteredInstancesColl?.Refresh();
    }
}
