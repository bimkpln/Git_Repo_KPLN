﻿using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Core.Reports;
using KPLN_Clashes_Ribbon.Services;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportWindow.xaml
    /// </summary>
    public partial class ReportWindow : Window
    {
        public List<IExecutableCommand> OnClosingActions = new List<IExecutableCommand>();

        private readonly Report _currentReport;
        private readonly SQLiteService_ReportItemsDB _sqliteService_ReportInstanceDB;
        private ObservableCollection<ReportItem> _reportInstancesColl;
        private readonly ReportManager _reportManager;
        private readonly SQLiteService_MainDB _sqliteService_MainDB = new SQLiteService_MainDB();
        private KPItemStatus _selectedKPItemStatus;
        private string _searchData = string.Empty;

        public ReportWindow(Report report, bool isEnabled, ReportManager reportManager)
        {
            _currentReport = report;
            _sqliteService_ReportInstanceDB = new SQLiteService_ReportItemsDB(_currentReport.PathToReportInstance);

            _reportManager = reportManager;

            _reportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();
            if (isEnabled)
            {
                foreach (ReportItem instance in _reportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Visible;
                    instance.IsControllsEnabled = true;
                }
            }
            else
            {
                foreach (ReportItem instance in _reportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Collapsed;
                    instance.IsControllsEnabled = false;
                }
            }

            InitializeComponent();

            Title = string.Format("KPLN: Отчет Navisworks ({0})", report.Name);

            UpdateCollection(KPItemStatus.Opened);

            Closing += RemoveOnClose;
        }

        /// <summary>
        /// Получить приоритетный статус из инстансов внутри одного отчета
        /// </summary>
        private KPItemStatus GetMainReportStatus()
        {
            _reportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();

            if (_reportInstancesColl.Any(c => c.Status == KPItemStatus.Opened))
                return KPItemStatus.Opened;
            else if (_reportInstancesColl.All(c => c.Status == KPItemStatus.Closed || c.Status == KPItemStatus.Approved))
                return KPItemStatus.Closed;
            else if (_reportInstancesColl.All(c => c.Status == KPItemStatus.Delegated))
                return KPItemStatus.Delegated;
            else if (_reportInstancesColl.All(c => c.Status == KPItemStatus.Delegated || c.Status == KPItemStatus.Approved || c.Status == KPItemStatus.Closed))
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
            ObservableCollection<ReportItem> filtered_collection = new ObservableCollection<ReportItem>();
            foreach (ReportItem report in _reportInstancesColl)
            {
                if (report.Element_1_Info.ToLower().Contains(_searchData) || report.Element_2_Info.ToLower().Contains(_searchData))
                    filtered_collection.Add(report);
            }
            ReportControll.ItemsSource = filtered_collection;
        }

        private void UpdateCollection(KPItemStatus status)
        {
            _selectedKPItemStatus = status;

            ObservableCollection<ReportItem> filtered_collection = new ObservableCollection<ReportItem>();
            foreach (ReportItem report in _reportInstancesColl)
            {
                if (_selectedKPItemStatus == KPItemStatus.Opened
                    && (report.Element_1_Info.ToLower().Contains(_searchData) || report.Element_2_Info.ToLower().Contains(_searchData)))
                {
                    if (report.Status == KPItemStatus.Opened || report.Status == KPItemStatus.Delegated)
                        filtered_collection.Add(report);
                }
                else if (report.Element_1_Info.ToLower().Contains(_searchData) || report.Element_2_Info.ToLower().Contains(_searchData))
                {
                    if (report.Status == _selectedKPItemStatus)
                        filtered_collection.Add(report);
                }

            }

            if (ReportControll != null)
                ReportControll.ItemsSource = filtered_collection;
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
                    UpdateCollection(KPItemStatus.Opened);
                    break;
                case 2:
                    UpdateCollection(KPItemStatus.Closed);
                    break;
                case 3:
                    UpdateCollection(KPItemStatus.Approved);
                    break;
                case 4:
                    UpdateCollection(KPItemStatus.Delegated);
                    break;
            }
        }

        private void UserInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = (System.Windows.Controls.TextBox)sender;
            _searchData = textBox.Text.ToLower();

            ObservableCollection<ReportItem> filtered_collection = new ObservableCollection<ReportItem>();
            foreach (ReportItem report in _reportInstancesColl)
            {
                if (_selectedKPItemStatus == KPItemStatus.Opened 
                    && (report.Element_1_Info.ToLower().Contains(_searchData) || report.Element_2_Info.ToLower().Contains(_searchData)))
                {
                    if (report.Status == KPItemStatus.Opened || report.Status == KPItemStatus.Delegated)
                        filtered_collection.Add(report);
                }
                else if (report.Element_1_Info.ToLower().Contains(_searchData) || report.Element_2_Info.ToLower().Contains(_searchData))
                {
                    if (report.Status == _selectedKPItemStatus)
                        filtered_collection.Add(report);
                }

            }

            if (ReportControll != null)
                ReportControll.ItemsSource = filtered_collection;
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
            _sqliteService_ReportInstanceDB.SetStatusId_ByReportItem(itemStatus, item);
            _sqliteService_ReportInstanceDB.SetComment_ByReportItem(msg, item);

            _sqliteService_MainDB.UpdateReportGroup_MarksLastChange_ByGroupId(_currentReport.ReportGroupId);
            _sqliteService_MainDB.UpdateReport_MarksLastChange_ByIdAndMainRepInstStatus(_currentReport.Id, GetMainReportStatus());

            item.Status = itemStatus;
            item.CommentCollection = ReportItemComment.ParseComments(_sqliteService_ReportInstanceDB.GetComment_ByReportItem(item), item);

            ResetDelegateBtnBrush(item);
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
            ReportItem item = subDepartmentBtn.Parent;
            if (subDepartmentBtn.Id == 7)
            {
                // Сброс выделения делегирования при нажатии на кнопку сброса (по id)
                ItemMessageWorker(item, KPItemStatus.Opened, $"Статус изменен: <Возвращен в работу - отказ в делегировании>\n");
            }
            else
            {
                // Выделение и логирования делегирования при нажатии на кнопку сброса (по id)
                SetDelegateBtnBrush(item, subDepartmentBtn);
                ItemMessageWorker(item, KPItemStatus.Delegated, $"Статус изменен: <Делегирована отделу {subDepartmentBtn.Name}>\n");
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
                { btn.DelegateBtnBackground = Brushes.Aqua; }
                else
                { sdBtn.DelegateBtnBackground = Brushes.Transparent; }

            }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            _reportManager.UpdateGroups();

            _reportInstancesColl = _sqliteService_ReportInstanceDB.GetAllReporItems();

            int i = cbxFilter.SelectedIndex;
            switch (i)
            {
                case 0:
                    UpdateCollection();
                    break;
                case 1:
                    UpdateCollection(KPItemStatus.Opened);
                    break;
                case 2:
                    UpdateCollection(KPItemStatus.Closed);
                    break;
                case 3:
                    UpdateCollection(KPItemStatus.Approved);
                    break;
                case 4:
                    UpdateCollection(KPItemStatus.Delegated);
                    break;
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
    }
}
