using System;
using System.Windows;

namespace KPLN_Library_Forms.UI
{
    /// <summary>Общее окно итогов операции. Не зависит от Revit и моделей вызывающего плагина.</summary>
    public partial class OperationReportWindow : Window
    {
        public string ReportText { get; }
        public string Summary { get; }
        public string ReportPath { get; }
        public string PathText => string.IsNullOrWhiteSpace(ReportPath) ? "" : "Отчёт: " + ReportPath;

        public OperationReportWindow(string title, string reportText, string summary = null, string reportPath = null)
        {
            ReportText = reportText ?? "";
            Summary = summary ?? "";
            ReportPath = reportPath ?? "";

            InitializeComponent();
            Title = string.IsNullOrWhiteSpace(title) ? "KPLN: Итоги выполнения" : title;
            SummaryBlock.Visibility = string.IsNullOrWhiteSpace(Summary) ? Visibility.Collapsed : Visibility.Visible;
            PathBlock.Visibility = string.IsNullOrWhiteSpace(ReportPath) ? Visibility.Collapsed : Visibility.Visible;
            DataContext = this;
        }

        private void CopyReport(object sender, RoutedEventArgs args)
        {
            try
            {
                string text = (string.IsNullOrWhiteSpace(Summary) ? "" : Summary + Environment.NewLine + Environment.NewLine)
                    + ReportText + (string.IsNullOrWhiteSpace(ReportPath) ? "" : Environment.NewLine + Environment.NewLine + PathText);
                Clipboard.SetText(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не удалось скопировать отчёт: " + ex.Message, Title);
            }
        }

        private void CloseReport(object sender, RoutedEventArgs args)
        {
            Close();
        }
    }
}
