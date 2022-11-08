using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Common;
using KPLN_Clashes_Ribbon.Common.Reports;
using KPLN_Clashes_Ribbon.Tools;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using static KPLN_Clashes_Ribbon.Common.Collections;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportWindow.xaml
    /// </summary>
    public partial class ReportWindow : Window
    {
        
        public List<IExecutableCommand> OnClosingActions = new List<IExecutableCommand>();

        private Report _currentReport;

        private ObservableCollection<ReportInstance> _reportInstancesColl;

        private ReportManager _reportManager;

        public ReportWindow(Report report, bool isEnabled, ReportManager reportManager)
        {
            _currentReport = report;

            _reportManager = reportManager;

            _reportInstancesColl = ReportInstance.GetReportInstances(report.Path);
            if (isEnabled)
            {
                foreach (ReportInstance instance in _reportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Visible;
                    instance.IsControllsEnabled = true;
                }
            }
            else
            {
                foreach (ReportInstance instance in _reportInstancesColl)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Collapsed;
                    instance.IsControllsEnabled = false;
                }
            }
            
            InitializeComponent();
            
            Title = string.Format("KPLN: Отчет Navisworks ({0})", report.Name);
            
            UpdateCollection(Status.Opened);
            
            Closing += RemoveOnClose;
        }

        /// <summary>
        /// Получить приоритетный статус из инстансов внутри одного отчета
        /// </summary>
        private Status GetMainReportStatus()
        {
            _reportInstancesColl = ReportInstance.GetReportInstances(_currentReport.Path);

            if (_reportInstancesColl.All(c => c.Status == Status.Closed || c.Status == Status.Approved))
            {
                return Status.Closed;
            }
            else if (_reportInstancesColl.All(c => c.Status == Status.Delegated))
            {
                return Status.Delegated;
            }
            else if (_reportInstancesColl.Any(c => c.Status == Status.Opened || c.Status == Status.Delegated))
            {
                return Status.Opened;
            }
            else
            {
                return Status.Opened;
            }
        }
        
        private void RemoveOnClose(object sender, CancelEventArgs args)
        {
            foreach (IExecutableCommand cmd in OnClosingActions)
            {
                try
                {
                    KPLN_Loader.Preferences.CommandQueue.Enqueue(cmd);
                }
                catch (Exception)
                { }
            }
        }
 
        private void OnLoadImage(object sender, RoutedEventArgs e)
        {
            ((sender as System.Windows.Controls.Button).DataContext as ReportInstance).LoadImage();
        }

        private void SelectId(object sender, RoutedEventArgs e)
        {
            try
            {
                int id = int.Parse((sender as System.Windows.Controls.Button).Content.ToString(), System.Globalization.NumberStyles.Integer);
                KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandZoomSelectElement(id));
            }
            catch (Exception)
            { }

        }

        private void PlacePoint(object sender, RoutedEventArgs args)
        {
            try
            {
                ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
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
                XYZ point = new XYZ(double.Parse(parts[0].Replace(".", ","), System.Globalization.NumberStyles.Float), double.Parse(parts[1].Replace(".", ","), System.Globalization.NumberStyles.Float), double.Parse(parts[2].Replace(".", ","), System.Globalization.NumberStyles.Float));
                KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandPlaceFamily(point, report.Element_1_Id, report.Element_1_Info, report.Element_2_Id, report.Element_2_Info, this));
            }
            catch (Exception)
            { }
        }

        private void UpdateCollection()
        {
            ObservableCollection<ReportInstance> filtered_collection = new ObservableCollection<ReportInstance>();
            foreach (ReportInstance report in _reportInstancesColl)
            {
                filtered_collection.Add(report);
            }
            ReportControll.ItemsSource = filtered_collection;
        }

        private void UpdateCollection(Status status)
        {
            try
            {
                ObservableCollection<ReportInstance> filtered_collection = new ObservableCollection<ReportInstance>();
                foreach (ReportInstance report in _reportInstancesColl)
                {
                    if (status == Status.Opened)
                    {
                        if (report.Status == Status.Opened || report.Status == Status.Delegated)
                        {
                            filtered_collection.Add(report);
                        }
                    }
                    else
                    {
                        if (report.Status == status)
                        {
                            filtered_collection.Add(report);
                        }
                    }
                    
                }
                if (ReportControll != null)
                { ReportControll.ItemsSource = filtered_collection; }
            }
            catch (Exception)
            { }
        }
        
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int i = this.cbxFilter.SelectedIndex;
                switch (i)
                {
                    case 0:
                        UpdateCollection();
                        break;
                    case 1:
                        UpdateCollection(Status.Opened);
                        break;
                    case 2:
                        UpdateCollection(Status.Closed);
                        break;
                    case 3:
                        UpdateCollection(Status.Approved);
                        break;
                    case 4:
                        UpdateCollection(Status.Delegated);
                        break;
                }
            }
            catch (Exception) { }
        }

        private void OnCorrected(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            try
            {
                DbController.SetInstanceValue(_currentReport.Path, report.Id, "STATUS", 0);
                DbController.SetInstanceValue(_currentReport.Path, report.Id, "DEPARTMENT", -1);
                report.Status = Common.Collections.Status.Closed;
                report.AddComment(string.Format("Статус изменен: <Исправлено>\n"), 1);
                DbController.UpdateGroupLastChange(_currentReport.GroupId);
                DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
            }
            catch (Exception ex)
            { PrintError(ex); }

            ResetDelegateBtnBrush(report);
        }

        private void OnApproved(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            TextInputDialog inputName = new TextInputDialog(this, "Введите комментарий:");
            inputName.ShowDialog();
            if (inputName.IsConfirmed())
            {
                try
                {
                    DbController.SetInstanceValue(_currentReport.Path, report.Id, "STATUS", 1);
                    report.Status = Common.Collections.Status.Approved;
                    report.AddComment(string.Format("Статус изменен: <Допустимое>\n" + inputName.GetLastPickedValue()), 1);
                    DbController.UpdateGroupLastChange(_currentReport.GroupId);
                    DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
                }
                catch (Exception ex)
                { PrintError(ex); }

                ResetDelegateBtnBrush(report);
            }
        }

        private void OnReset(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            try
            {
                DbController.SetInstanceValue(_currentReport.Path, report.Id, "STATUS", -1);
                report.Status = Common.Collections.Status.Opened;
                report.AddComment(string.Format("Статус изменен: <Открытое>\n"), 1);
                DbController.UpdateGroupLastChange(_currentReport.GroupId);
                DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
            }
            catch (Exception ex)
            { PrintError(ex); }
        }

        /// <summary>
        /// Кнопка делегирования коллизии на другой отдел
        /// </summary>
        private void OnDelegate(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            SubDepartmentBtn subDepartmentBtn = button.DataContext as SubDepartmentBtn;
            ReportInstance report = subDepartmentBtn.Parent;
            try
            {
                if (subDepartmentBtn.Id == 7)
                {
                    // Сброс выделения делегирования при нажатии на кнопку сброса (по id)
                    DbController.SetInstanceValue(_currentReport.Path, report.Id, "STATUS", -1);
                    DbController.SetInstanceValue(_currentReport.Path, report.Id, "DEPARTMENT", -1);
                    report.Status = Common.Collections.Status.Opened;
                    report.AddComment(string.Format($"Статус изменен: <Возвращен в работу>\n"), 1);
                    DbController.UpdateGroupLastChange(_currentReport.GroupId);
                    DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
                    ResetDelegateBtnBrush(report);
                }
                else 
                {
                    // Выделение и логирования делегирования при нажатии на кнопку сброса (по id)
                    DbController.SetInstanceValue(_currentReport.Path, report.Id, "STATUS", 2);
                    DbController.SetInstanceValue(_currentReport.Path, report.Id, "DEPARTMENT", subDepartmentBtn.Id);
                    report.Status = Common.Collections.Status.Delegated;
                    report.AddComment(string.Format($"Статус изменен: <Делегирована отделу {subDepartmentBtn.Name}>\n"), 1);
                    DbController.UpdateGroupLastChange(_currentReport.GroupId);
                    DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
                    SetDelegateBtnBrush(report, subDepartmentBtn);
                }
                
            }
            catch (Exception ex)
            { PrintError(ex); }
        }

        /// <summary>
        /// Сброс цвета кнопок делегирования
        /// </summary>
        private void ResetDelegateBtnBrush(ReportInstance report)
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
        private void SetDelegateBtnBrush(ReportInstance report, SubDepartmentBtn btn)
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

        private void OnAddComment(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            TextInputDialog inputName = new TextInputDialog(this, "Введите комментарий:");
            inputName.ShowDialog();
            if (inputName.IsConfirmed())
            {
                report.AddComment(inputName.GetLastPickedValue(), 0);
                DbController.UpdateGroupLastChange(_currentReport.GroupId);
                DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
            }
        }

        private void OnRemoveComment(object sender, RoutedEventArgs e)
        {
            try
            {
                ReportComment comment = (sender as System.Windows.Controls.Button).DataContext as ReportComment;
                comment.Parent.RemoveComment(comment);
                try
                {
                    DbController.UpdateGroupLastChange(_currentReport.GroupId);
                    DbController.UpdateReportLastChange(_currentReport.Id, GetMainReportStatus());
                }
                catch (Exception)
                { }
            }
            catch (Exception)
            { }
        }

        private void OnBtnUpdate(object sender, RoutedEventArgs args)
        {
            _reportManager.UpdateGroups();

            _reportInstancesColl = ReportInstance.GetReportInstances(_currentReport.Path);
            
            int i = cbxFilter.SelectedIndex;
            switch (i)
            {
                case 0:
                    UpdateCollection();
                    break;
                case 1:
                    UpdateCollection(Status.Opened);
                    break;
                case 2:
                    UpdateCollection(Status.Closed);
                    break;
                case 3:
                    UpdateCollection(Status.Approved);
                    break;
                case 4:
                    UpdateCollection(Status.Delegated);
                    break;
            }
        }
       private void OnExport(object sender, RoutedEventArgs e)
        {
            List<string> rows = new List<string>();
            ObservableCollection<ReportInstance> visibleCollection = ReportControll.ItemsSource as ObservableCollection<ReportInstance>;
            foreach (ReportInstance instance in visibleCollection)
            {
                rows.Add(GetExportRow(instance));
                foreach (ReportInstance subInstance in instance.SubElements)
                {
                    rows.Add(GetExportRow(subInstance, instance));
                }
            }
            string data = string.Join(Environment.NewLine, rows);
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Сохранить отчет";
            dialog.Filter = "txt file (*.txt)|*.txt";
            dialog.RestoreDirectory = true;
            dialog.FileName = string.Format("{0}_{1}-{2}-{3}.txt", _currentReport.Name.Replace('.', '_'), DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString());
            dialog.CreatePrompt = false;
            dialog.OverwritePrompt = true;
            dialog.SupportMultiDottedExtensions = false;
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
        
        private static string GetExportRow(ReportInstance instance, ReportInstance parent = null)
        {

            List<string> parts = new List<string>();
            parts.Add(Optimize(instance.Id.ToString()));
            parts.Add(Optimize(instance.GroupId.ToString()));
            parts.Add(Optimize(instance.Name));
            if (parent != null)
            {
                parts.Add(Optimize(parent.Status.ToString("G")));
                if (parent.Comments.Count != 0)
                {
                    ReportComment comment = parent.Comments.Last();
                    parts.Add(string.Format("<{0}> {1} {2}", Optimize(comment.Time), Optimize(comment.User), Optimize(comment.Message)));
                }
                else
                {
                    parts.Add(Optimize("..."));
                }
            }
            else
            {
                parts.Add(Optimize(instance.Status.ToString("G")));
                if (instance.Comments.Count != 0)
                {
                    ReportComment comment = instance.Comments.Last();
                    parts.Add(string.Format("<{0}> {1} {2}", Optimize(comment.Time), Optimize(comment.User), Optimize(comment.Message)));
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
