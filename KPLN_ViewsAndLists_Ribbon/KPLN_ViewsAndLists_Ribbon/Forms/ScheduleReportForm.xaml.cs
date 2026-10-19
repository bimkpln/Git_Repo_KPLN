using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using System;
using System.Windows;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ScheduleReportForm : Window
    {
        private readonly ScheduleExecutionReport _report;

        public ScheduleReportForm(ScheduleExecutionReport report)
        {
            _report = report;
            InitializeComponent();
            DataContext = report;
        }

        private void CopyReport(object sender, RoutedEventArgs args)
        {
            try
            {
                Clipboard.SetText(_report.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не удалось скопировать отчёт: " + ex.Message, Title);
            }
        }
    }
}
