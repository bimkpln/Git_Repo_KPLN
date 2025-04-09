using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Library_Bitrix24Worker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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

        private readonly ReportManagerForm _reportManager;
        private readonly ReportGroup _repourtGroup;
        private readonly Report _currentReport;
        private readonly Services.SQLite.SQLiteService_ReportItemsDB _sqliteService_ReportInstanceDB;
        private readonly Services.SQLite.SQLiteService_MainDB _sqliteService_MainDB = new Services.SQLite.SQLiteService_MainDB();

        private string _conflictDataTBx = string.Empty;
        private string _idDataTBx = string.Empty;
        private string _conflictMetaDataTBx = string.Empty;

        public ReportForm(Report report, ReportGroup reportGroup, ReportManagerForm reportManager)
        {
            _reportManager = reportManager;
            _repourtGroup = reportGroup;
            _currentReport = report;

            _sqliteService_ReportInstanceDB = new Services.SQLite.SQLiteService_ReportItemsDB(_currentReport.PathToReportInstance);

            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();
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

            Closing += RemoveOnClose;
        }

        /// <summary>
        /// Исходная коллекция элементов
        /// </summary>
        public ObservableCollection<ReportItem> ReportInstancesColl { get; private set; }

        /// <summary>
        /// Отфильтрованная коллекция элементов, которая является контекстом для окна
        /// </summary>
        public ObservableCollection<ReportItem> FilteredInstancesColl { get; private set; } = new ObservableCollection<ReportItem>();

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
            {
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandZoomSelectElement(id, elInfo));
            }
            else
                throw new Exception("Проблемы с отчетом: параметр id не парсится");
        }

        private void PlacePoint(object sender, RoutedEventArgs args)
        {
            ReportItem report = (sender as System.Windows.Controls.Button).DataContext as ReportItem;

            string pt = report.Point;
            pt = pt.Replace("X:", "");
            pt = pt.Replace("Y:", "");
            pt = pt.Replace("Z:", "");

            string pts = string.Empty;
            foreach (char c in pt)
            {
                if ("-0123456789.,".Contains(c))
                {
                    pts += c;
                }
            }

            string[] parts = pts.Split(',');
            //var temp = double.Parse(parts[0].Replace(".", ","), NumberStyles.Float);
            if (
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointX)
                && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointY)
                && double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double pointZ)
                )
            {
                XYZ point = new XYZ(pointX, pointY, pointZ);
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandPlaceFamily(point, report.Element_1_Id, report.Element_1_Info, report.Element_2_Id, report.Element_2_Info, this));
            }
            else
                throw new Exception("Проблемы с CultureInfo");
        }

        private void UpdateCollection()
        {
            foreach (ReportItem report in ReportInstancesColl)
            {
                AddToFilteredInstancesColl_ByUserFilterData(report);
            }
        }

        private void UpdateCollection_ByStatys(KPItemStatus status)
        {
            foreach (ReportItem report in ReportInstancesColl)
            {
                if (report.Status == status)
                    AddToFilteredInstancesColl_ByUserFilterData(report);
                else
                    FilteredInstancesColl.Remove(report);
            }
        }

        private void UpdateCollection_ByStatuses(KPItemStatus status1, KPItemStatus status2)
        {
            foreach (ReportItem report in ReportInstancesColl)
            {
                if (report.Status == status1 || report.Status == status2)
                    AddToFilteredInstancesColl_ByUserFilterData(report);
                else
                    FilteredInstancesColl.Remove(report);
            }
        }

        /// <summary>
        /// Главный метод фильтарции отчетов по статусу и пользовательским полям
        /// </summary>
        /// <param name="report"></param>
        private void AddToFilteredInstancesColl_ByUserFilterData(ReportItem report)
        {
            bool isEmptyConflData = string.IsNullOrEmpty(_conflictDataTBx);
            bool isEmptyConflictMetaData = string.IsNullOrEmpty(_conflictMetaDataTBx);
            bool isEmptyIDData = string.IsNullOrEmpty(_idDataTBx);

            if (isEmptyConflData && isEmptyConflictMetaData && isEmptyIDData)
            {
                if (!FilteredInstancesColl.Contains(report))
                    FilteredInstancesColl.Add(report);
            }
            else
            {
                bool checkConflData = !isEmptyConflData && report.Name.IndexOf(_conflictDataTBx, StringComparison.OrdinalIgnoreCase) >= 0;
                bool checkConflictMetaData = !isEmptyConflictMetaData
                    && (report.Element_1_Info.IndexOf(_conflictMetaDataTBx, StringComparison.OrdinalIgnoreCase) >= 0
                        || report.Element_2_Info.IndexOf(_conflictMetaDataTBx, StringComparison.OrdinalIgnoreCase) >= 0);
                bool checkIDData = !isEmptyIDData
                    && (report.Element_1_Id.ToString().IndexOf(_idDataTBx, StringComparison.OrdinalIgnoreCase) >= 0
                        || report.Element_2_Id.ToString().IndexOf(_idDataTBx, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (report.SubElements.Select(se => se.Element_1_Id.ToString()).Contains(_idDataTBx)
                        || report.SubElements.Select(se => se.Element_2_Id.ToString()).Contains(_idDataTBx));

                if (checkConflData || checkConflictMetaData || checkIDData)
                {
                    if (!FilteredInstancesColl.Contains(report))
                        FilteredInstancesColl.Add(report);
                }
                else
                    FilteredInstancesColl.Remove(report);
            }
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int i = this.cbxFilter.SelectedIndex;
            switch (i)
            {
                case 0:
                    UpdateCollection();
                    break;
                case 1:
                    UpdateCollection_ByStatuses(KPItemStatus.Opened, KPItemStatus.Delegated);
                    break;
                case 2:
                    UpdateCollection_ByStatys(KPItemStatus.Closed);
                    break;
                case 3:
                    UpdateCollection_ByStatys(KPItemStatus.Approved);
                    break;
                case 4:
                    UpdateCollection_ByStatys(KPItemStatus.Delegated);
                    break;
            }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            _reportManager.UpdateGroups();

            ReportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();

            OnSelectionChanged(null, null);
        }

        private void ConflictDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            OnSelectionChanged(null, null);

            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _conflictDataTBx = textBox.Text;

            UpdateCollection();

            OnSelectionChanged(null, null);
        }

        private void IDDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            OnSelectionChanged(null, null);

            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _idDataTBx = textBox.Text;

            UpdateCollection();

            OnSelectionChanged(null, null);
        }

        private void ConflictMetaDataTBx_TextChanged(object sender, TextChangedEventArgs e)
        {
            OnSelectionChanged(null, null);

            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _conflictMetaDataTBx = textBox.Text;

            UpdateCollection();

            OnSelectionChanged(null, null);
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

        private void OnCorrected(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            ItemMessageWorker(item, KPItemStatus.Closed, "Статус изменен: <Устранено>\n");
        }

        private void OnApproved(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;

            TextInputForm textInputForm = new TextInputForm(this, "Введите комментарий:");
            textInputForm.ShowDialog();
            string msg = textInputForm.UserComment;
            if (msg != null)
                ItemMessageWorker(item, KPItemStatus.Approved, $"Статус изменен: <Допустимое>\n{msg}");
        }

        private void OnReset(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;
            ItemMessageWorker(item, KPItemStatus.Opened, $"Статус изменен: <Возвращен в работу - отказ в допуске>\n");
        }

        private void OnAddComment(object sender, RoutedEventArgs e)
        {
            ReportItem item = (sender as System.Windows.Controls.Button).DataContext as ReportItem;

            TextInputForm textInputForm = new TextInputForm(this, "Введите комментарий:");
            textInputForm.ShowDialog();
            string msg = textInputForm.UserComment;
            if (msg != null)
                ItemMessageWorker(item, $"{msg}");
        }

        /// <summary>
        /// Кнопка делегирования коллизии на другой отдел
        /// </summary>
        private void OnDelegate(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            SubDepartmentBtn subDepartmentBtn = button.DataContext as SubDepartmentBtn;
            if (!(FindParent<ItemsControl>(button).DataContext is ReportItem item))
                return;

            // Сброс выделения делегирования при нажатии на кнопку сброса (по id)
            if (subDepartmentBtn.Id == 99)
            {
                ResetDelegateBtnBrush(item);
                ItemMessageWorker(item, KPItemStatus.Opened, $"Статус изменен: <Возвращен в работу - отказ в делегировании>\n");
            }
            // Выделение и логирования делегирования при нажатии на кнопку сброса (по id)
            else
            {
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
                    Task<bool> sendMsgToTaskTask = Task<bool>.Run(() =>
                    {
                        return BitrixMessageSender
                            .SendMsgToTask_ByTaskId(
                            delegBitrixTaskId, 
                            $"Пользователь <{CurrentDBUser.Name} {CurrentDBUser.Surname}> делегировал вам коллизию из отчета: \"{_currentReport.Name}\"");
                    });

                    if (sendMsgToTaskTask.Result)
                        System.Windows.MessageBox.Show(
                            $"Было отправлено сообщение о делегировании коллизии в задачу Bitrix с id: {delegBitrixTaskId}", 
                            "Bitrix", 
                            (MessageBoxButton)MessageBoxButtons.OK, 
                            (MessageBoxImage)MessageBoxIcon.Asterisk);
                }

                SetDelegateBtnBrush(item, subDepartmentBtn);
                ItemMessageWorker(item, KPItemStatus.Delegated, $"Статус изменен: <Делегирована отделу {subDepartmentBtn.Name}>\n");
            }

            // Обновление коллекции по делегировнным кнопкам
            foreach (ReportItem ri in ReportInstancesColl)
            {
                if (ri.Id == item.Id)
                    ri.SubDepartmentBtns = item.SubDepartmentBtns;
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
