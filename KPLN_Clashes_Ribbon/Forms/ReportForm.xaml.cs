using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker;
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
        public List<IExecutableCommand> OnClosingActions = new List<IExecutableCommand>();

        private readonly ReportGroup _repourtGroup;
        private readonly Report _currentReport;
        private readonly Services.SQLite.SQLiteService_ReportItemsDB _sqliteService_ReportInstanceDB;
        private readonly Services.SQLite.SQLiteService_MainDB _sqliteService_MainDB = new Services.SQLite.SQLiteService_MainDB();

        private string _conflictDataTBx = string.Empty;
        private string _idDataTBx = string.Empty;
        private string _conflictMetaDataTBx = string.Empty;

        public ReportForm(Report report, ReportGroup reportGroup)
        {
            _repourtGroup = reportGroup;
            _currentReport = report;

            _sqliteService_ReportInstanceDB = new Services.SQLite.SQLiteService_ReportItemsDB(_currentReport.PathToReportInstance);

            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();
            FilteredInstancesColl = CollectionViewSource.GetDefaultView(ReportInstancesColl);
            FilteredInstancesColl.Filter += FilterForRepItems;
            
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

            Title = string.Format("KPLN: Отчет Navisworks ({0})", report.Name);
            DataContext = this;

            OnClosingActions.Add(new CommandRemoveInstance());
            Closing += RemoveOnClose;
        }

        private bool FilterForRepItems(object obj)
        {
            if (obj is ReportItem item)
            {
                System.Windows.Controls.ComboBox sfCBX = this.StatusFilterCBX;
                if (sfCBX == null)
                    return true;

                int mainStatusFilterIndex = this.StatusFilterCBX.SelectedIndex;
                // Фильтрация по статусу
                switch (mainStatusFilterIndex)
                {
                    case 0:
                        return CheckParamData(item);
                    case 1:
                        if (item.Status == KPItemStatus.Opened || item.Status == KPItemStatus.Delegated)
                            return CheckParamData(item);

                        return false;
                    case 2:
                        if (item.Status == KPItemStatus.Closed)
                            return CheckParamData(item);

                        return false;
                    case 3:
                        if (item.Status == KPItemStatus.Approved)
                            return CheckParamData(item);

                        return false;
                    case 4:
                        if (item.Status == KPItemStatus.Delegated)
                            return CheckParamData(item);

                        return false;
                }
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

                bool checkIDData = item.GroupElementIds.Contains(_idDataTBx);
                    
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

        /// <summary>
        /// Получить приоритетный статус из инстансов внутри одного отчета
        /// </summary>
        private KPItemStatus GetMainReportStatus()
        {
            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();

            if (ReportInstancesColl.Any(c => c.Status == KPItemStatus.Opened))
                return KPItemStatus.Opened;
            else if (ReportInstancesColl.All(c => c.Status == KPItemStatus.Closed || c.Status == KPItemStatus.Approved))
                return KPItemStatus.Closed;
            else if (ReportInstancesColl.All(c => c.Status == KPItemStatus.Delegated))
                return KPItemStatus.Delegated;
            else if (ReportInstancesColl.All(c => c.Status == KPItemStatus.Delegated || c.Status == KPItemStatus.Approved || c.Status == KPItemStatus.Closed))
                return KPItemStatus.Delegated;
            else
                return KPItemStatus.Opened;
        }

        private void RemoveOnClose(object sender, CancelEventArgs args)
        {
            foreach (IExecutableCommand cmd in OnClosingActions)
            {
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(cmd);
            }
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

        private void SelectId(object sender, string elInfo)
        {
            if (int.TryParse((sender as System.Windows.Controls.Button).Content.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandZoomSelectElement(id, elInfo));
            else
                throw new Exception("Проблемы с отчетом: параметр id не парсится");
        }

        private void PlacePoint(object sender, RoutedEventArgs args)
        {
            if ((sender as System.Windows.Controls.Button).DataContext is ReportItem report)
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandPlaceFamily(report));
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

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();

            FilteredInstancesColl?.Refresh();
        }

        /// <summary>
        /// Метод для записи данных (комментарии) в окно и БД для текущего ReportItem
        /// </summary>
        /// <param name="item">ReportItem для анализа</param>
        /// <param name="msg">Сообщение при смене статуса</param>
        private void ItemMessageWorker(ReportItem item, string msg)
        {
            _sqliteService_ReportInstanceDB.SetComment_ByReportItem(msg, item);

            _sqliteService_MainDB.UpdateReportGroup_MarksLastChange_ByGroupId(_currentReport.ReportGroupId);
            _sqliteService_MainDB.UpdateReport_MarksLastChange_ByIdAndMainRepInstStatus(_currentReport.Id, GetMainReportStatus());

            item.CommentCollection = ReportItemComment.ParseComments(_sqliteService_ReportInstanceDB.GetComment_ByReportItem(item), item);
        }

        /// <summary>
        /// Метод для записи данных (комментарии, статус) в окно и БД для текущего ReportItem
        /// </summary>
        /// <param name="item">ReportItem для анализа</param>
        /// <param name="itemStatus">Присаваиваемый статус замечания</param>
        /// <param name="msg">Сообщение при смене статуса</param>
        private void ItemMessageWorker(ReportItem item, KPItemStatus itemStatus, string msg)
        {
            _sqliteService_ReportInstanceDB.SetStatusAndDepartment_ByReportItem(itemStatus, item);
            _sqliteService_ReportInstanceDB.SetComment_ByReportItem(msg, item);

            _sqliteService_MainDB.UpdateReportGroup_MarksLastChange_ByGroupId(_currentReport.ReportGroupId);
            _sqliteService_MainDB.UpdateReport_MarksLastChange_ByIdAndMainRepInstStatus(_currentReport.Id, GetMainReportStatus());

            item.Status = itemStatus;
            item.CommentCollection = ReportItemComment.ParseComments(_sqliteService_ReportInstanceDB.GetComment_ByReportItem(item), item);
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

                TaskDialogResult tdResult = td.Show();
                if (tdResult == TaskDialogResult.No)
                    return new List<ReportItem>();
            }

            return items;
        }

        private void OnCorrected(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            foreach (ReportItem ri in GetTargetItems(item))
            {
                ItemMessageWorker(ri, KPItemStatus.Closed, "Статус изменен: <Устранено>\\n");
            }
        }

        private void OnApproved(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;

            List<ReportItem> items = GetTargetItems(item);

            TextInputForm textInputForm = new TextInputForm(this, "Введите комментарий:");
            textInputForm.ShowDialog();
            string msg = textInputForm.UserComment;
            if (msg == null)
                return;

            foreach (ReportItem ri in items)
            {
                ItemMessageWorker(ri, KPItemStatus.Approved, $"Статус изменен: <Допустимое>\n{msg}");
            }
        }

        private void OnReset(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            foreach (ReportItem ri in GetTargetItems(item))
            {
                ItemMessageWorker(ri, KPItemStatus.Opened, $"Статус изменен: <Возвращен в работу - отказ в допуске>\\n");
            }
        }

        private void OnAddComment(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;

            TextInputForm textInputForm = new TextInputForm(this, "Введите комментарий:");
            textInputForm.ShowDialog();
            string msg = textInputForm.UserComment;
            if (msg == null)
                return;

            foreach (ReportItem ri in GetTargetItems(item))
            {
                ItemMessageWorker(ri, $"{msg}");
            }
        }

        /// <summary>
        /// Кнопка делегирования коллизии на другой отдел
        /// </summary>
        private void OnDelegate(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            SubDepartmentBtn subDepartmentBtn = button.DataContext as SubDepartmentBtn;
            if (!(FindParent<ItemsControl>(button).DataContext is ReportItem currentItem))
                return;

            List<ReportItem> items = GetTargetItems(currentItem);

            // Сброс выделения делегирования при нажатии на кнопку сброса (по id)
            if (subDepartmentBtn.Id == 99)
            {
                foreach (ReportItem ri in items)
                {
                    ResetDelegateBtnBrush(ri);
                    ItemMessageWorker(ri, KPItemStatus.Opened, $"Статус изменен: <Возвращен в работу - отказ в делегировании>\n");
                    foreach (ReportItem r in ReportInstancesColl)
                    {
                        if (r.Id == ri.Id)
                            r.SubDepartmentBtns = ri.SubDepartmentBtns;
                    }
                }

                return;
            }

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

            // Анализ задачи в Bitrix (если есть)
            if (delegBitrixTaskId > 0)
            {
                Task<bool> isTaskOpenTask = Task<bool>.Run(() =>
                    BitrixMessageSender.CheckTaskOpens_ByTaskId(delegBitrixTaskId));

                // Спам ТОЛЬКО в закрытые задачи
                if (!isTaskOpenTask.Result)
                {
                    Task<bool> sendMsgToTaskTask = Task<bool>.Run(() =>
                        BitrixMessageSender
                            .SendMsgToTask_ByTaskId(
                            delegBitrixTaskId,
                            $"Пользователь <{DBMainService.CurrentDBUser.Name} {DBMainService.CurrentDBUser.Surname}> делегировал вам коллизию из отчета: \"{_currentReport.Name}\""));

                    if (sendMsgToTaskTask.Result)
                        System.Windows.MessageBox.Show(
                            $"ВНИМАНИЕ! Отдел, которому вы делегируете замечание - уже отработал свою задачу. Свяжитесь с исполнителем лично" +
                            $"\nИНФО: Было отправлено сообщение о делегировании коллизии в задачу Bitrix с id: {delegBitrixTaskId}",
                            "Bitrix",
                            (MessageBoxButton)MessageBoxButtons.OK,
                            (MessageBoxImage)MessageBoxIcon.Asterisk);
                }
            }

            foreach (ReportItem ri in items)
            {
                SubDepartmentBtn targetBtn = ri.SubDepartmentBtns.FirstOrDefault(b => b.Id == subDepartmentBtn.Id);
                if (targetBtn != null)
                    SetDelegateBtnBrush(ri, targetBtn);

                ItemMessageWorker(ri, KPItemStatus.Delegated, $"Статус изменен: <Делегирована отделу {subDepartmentBtn.Name}>\n");

                // Обновление коллекции по делегировнным кнопкам
                foreach (ReportItem r in ReportInstancesColl)
                {
                    if (r.Id == ri.Id)
                        r.SubDepartmentBtns = ri.SubDepartmentBtns;
                }
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
            ObservableCollection<ReportItem> visibleCollection = ReportControll.ItemsSource as ObservableCollection<ReportItem>;
            foreach (ReportItem instance in visibleCollection)
            {
                rows.Add(GetExportRow(instance));
                foreach (ReportItem subInstance in instance.SubElements)
                {
                    rows.Add(GetExportRow(subInstance, instance));
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
            parts.Add(Optimize(instance.Element_1_Id.ToString()));
            parts.Add(Optimize(instance.Element_1_Info));
            parts.Add(Optimize(instance.Element_2_Id.ToString()));
            parts.Add(Optimize(instance.Element_2_Info));
            parts.Add(Optimize(instance.Point));
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
    }
}
