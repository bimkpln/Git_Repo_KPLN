using Autodesk.Revit.DB;
using KPLN_Loader.Common;
using KPLN_Clashes_Ribbon.Commands;
using KPLN_Clashes_Ribbon.Common.Reports;
using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Loader.Output.Output;
using static KPLN_Clashes_Ribbon.Common.Collections;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ReportWindow.xaml
    /// </summary>
    public partial class ReportWindow : Window
    {
        public List<IExecutableCommand> OnClosingActions = new List<IExecutableCommand>();
        private Report Report { get; }
        private ObservableCollection<ReportInstance> Collection { get; set; }
        private bool AllElementsClosed()
        {
            try
            {
                foreach (ReportInstance report in Collection)
                {
                    if (report.Status == Status.Opened)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }
        public ReportWindow(Report report, ObservableCollection<ReportInstance> reports, bool isEnabled)
        {
            if (isEnabled)
            {
                foreach (ReportInstance instance in reports)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Visible;
                    instance.IsControllsEnabled = true;
                }
            }
            else
            {
                foreach (ReportInstance instance in reports)
                {
                    instance.IsControllsVisible = System.Windows.Visibility.Collapsed;
                    instance.IsControllsEnabled = false;
                }
            }
            Collection = reports;
            Report = report;
            InitializeComponent();
            Title = string.Format("KPLN: Отчет Navisworks ({0})", report.Name);
            //ReportControll.ItemsSource = Collection;
            UpdateCollection(Status.Opened);
            Closing += RemoveOnClose;
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
        private void OnCorrected(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            TextInputDialog inputName = new TextInputDialog(this, "Введите комментарий:");
            inputName.ShowDialog();
            if (inputName.IsConfirmed())
            {
                try
                {
                    DbController.SetInstanceValue(Report.Path, report.Id, "STATUS", 0);
                    report.Status = Common.Collections.Status.Closed;
                    DbController.UpdateGroupLastChange(Report.GroupId);
                    DbController.UpdateReportLastChange(Report.Id, AllElementsClosed());
                    report.AddComment(string.Format("Статус изменен: <Исправлено>\n" + inputName.GetLastPickedValue()), 1);
                }
                catch (Exception ex)
                { PrintError(ex); }
            }
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
                    DbController.SetInstanceValue(Report.Path, report.Id, "STATUS", 1);
                    report.Status = Common.Collections.Status.Approved;
                    DbController.UpdateGroupLastChange(Report.GroupId);
                    DbController.UpdateReportLastChange(Report.Id, AllElementsClosed());
                    report.AddComment(string.Format("Статус изменен: <Допустимое>\n" + inputName.GetLastPickedValue()), 1);
                }
                catch (Exception ex)
                { PrintError(ex); }
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
        private void UpdateCollection(Status status)
        {
            try
            {
                ObservableCollection<ReportInstance> filtered_collection = new ObservableCollection<ReportInstance>();
                foreach (ReportInstance report in Collection)
                {
                    if (report.Status == status)
                    {
                        filtered_collection.Add(report);
                    }
                }
                ReportControll.ItemsSource = null;
                ReportControll.ItemsSource = filtered_collection;
            }
            catch (Exception)
            { }
        }
        private void UpdateCollection()
        {
            ObservableCollection<ReportInstance> filtered_collection = new ObservableCollection<ReportInstance>();
            foreach (ReportInstance report in Collection)
            {
                filtered_collection.Add(report);
            }
            ReportControll.ItemsSource = filtered_collection;
        }
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                int i = this.cbxFilter.SelectedIndex;
                if (i == 0)
                {
                    UpdateCollection();
                }
                if (i == 1)
                {
                    UpdateCollection(Status.Opened);
                }
                if (i == 2)
                {
                    UpdateCollection(Status.Closed);
                }
                if (i == 3)
                {
                    UpdateCollection(Status.Approved);
                }
            }
            catch (Exception){ }
        }

        private void OnReset(object sender, RoutedEventArgs e)
        {
            ReportInstance report = (sender as System.Windows.Controls.Button).DataContext as ReportInstance;
            TextInputDialog inputName = new TextInputDialog(this, "Введите комментарий:");
            inputName.ShowDialog();
            if (inputName.IsConfirmed())
            {
                try
                {
                    DbController.SetInstanceValue(Report.Path, report.Id, "STATUS", -1);
                    report.Status = Common.Collections.Status.Opened;
                    DbController.UpdateGroupLastChange(Report.GroupId);
                    DbController.UpdateReportLastChange(Report.Id, AllElementsClosed());
                    report.AddComment(string.Format("Статус изменен: <Открытое>\n" + inputName.GetLastPickedValue()), 1);
                }
                catch (Exception ex)
                { PrintError(ex); }
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
                DbController.UpdateGroupLastChange(Report.GroupId);
                DbController.UpdateReportLastChange(Report.Id, AllElementsClosed());
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
                    DbController.UpdateGroupLastChange(Report.GroupId);
                    DbController.UpdateReportLastChange(Report.Id, AllElementsClosed());
                }
                catch (Exception)
                { }
            }
            catch (Exception)
            { }
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
            dialog.FileName = string.Format("{0}_{1}-{2}-{3}.txt", Report.Name.Replace('.', '_'), DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString());
            dialog.CreatePrompt = false;
            dialog.OverwritePrompt = true;
            dialog.SupportMultiDottedExtensions = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(dialog.OpenFile());
                foreach(string row in rows)
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
