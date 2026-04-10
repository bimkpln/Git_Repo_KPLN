using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace KPLN_Tools.Forms
{
    public partial class ApartmentExecutionReportWindow : Window
    {
        private readonly List<ApartmentExecutionReportItem> _items;

        public ApartmentExecutionReportWindow(List<ApartmentExecutionReportItem> items)
        {
            InitializeComponent();

            _items = items ?? new List<ApartmentExecutionReportItem>();
            DataContext = _items;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SaveReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Title = "Сохранить отчёт";
                dlg.Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dlg.FileName = "Отчёт_построения_квартир_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt";
                dlg.DefaultExt = ".txt";
                dlg.AddExtension = true;

                bool? result = dlg.ShowDialog(this);
                if (result != true)
                    return;

                string reportText = BuildPlainTextReport(_items);
                File.WriteAllText(dlg.FileName, reportText, Encoding.UTF8);

                MessageBox.Show(
                    this,
                    "Отчёт успешно сохранён:\n" + dlg.FileName,
                    "Сохранение отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "Не удалось сохранить отчёт:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static string BuildPlainTextReport(List<ApartmentExecutionReportItem> items)
        {
            StringBuilder sb = new StringBuilder();

            if (items == null || items.Count == 0)
            {
                sb.AppendLine("Отчёт пуст.");
                return sb.ToString();
            }

            foreach (ApartmentExecutionReportItem item in items)
            {
                if (item == null)
                    continue;

                sb.AppendLine(item.HeaderText);

                if (item.Lines != null && item.Lines.Count > 0)
                {
                    foreach (ApartmentExecutionReportLine line in item.Lines)
                    {
                        if (line == null || string.IsNullOrWhiteSpace(line.Text))
                            continue;

                        sb.AppendLine(line.Text);
                    }
                }

                sb.AppendLine();
            }

            return sb.ToString();
        }
    }

    public class ApartmentExecutionReportItem
    {
        public long ApartmentId { get; set; }
        public string CustomHeaderText { get; set; }

        public string HeaderText
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(CustomHeaderText))
                    return CustomHeaderText;

                return "2D СЕМЕЙСТВО КВАРТИРЫ [" + ApartmentId + "]";
            }
        }

        public ObservableCollection<ApartmentExecutionReportLine> Lines { get; set; }

        public ApartmentExecutionReportItem()
        {
            Lines = new ObservableCollection<ApartmentExecutionReportLine>();
        }
    }

    public class ApartmentExecutionReportLine
    {
        public string Text { get; set; }
        public Brush Foreground { get; set; }
        public double FontSize { get; set; }
        public FontWeight FontWeight { get; set; }

        public ApartmentExecutionReportLine()
        {
            Foreground = Brushes.Black;
            FontSize = 13;
            FontWeight = FontWeights.Normal;
        }
    }
}