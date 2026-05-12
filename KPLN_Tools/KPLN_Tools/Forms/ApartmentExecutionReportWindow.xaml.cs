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
    public interface IApartmentExecutionReportActionController
    {
        void RequestShowElement(ApartmentExecutionReportItem item);
        void RequestDeleteRemnants(ApartmentExecutionReportItem item);
        void RequestRestore2DFamily(ApartmentExecutionReportItem item);
    }

    public partial class ApartmentExecutionReportWindow : Window
    {
        private readonly List<ApartmentExecutionReportItem> _items;
        private readonly IApartmentExecutionReportActionController _actionController;

        public ApartmentExecutionReportWindow(List<ApartmentExecutionReportItem> items)
            : this(items, null)
        {
        }

        public ApartmentExecutionReportWindow(List<ApartmentExecutionReportItem> items, IApartmentExecutionReportActionController actionController)
        {
            InitializeComponent();

            _items = items ?? new List<ApartmentExecutionReportItem>();
            _actionController = actionController;
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

        private void NavigateToElement_Click(object sender, RoutedEventArgs e)
        {
            ApartmentExecutionReportItem item = GetReportItem(sender);
            if (item == null || _actionController == null)
                return;

            WindowState = WindowState.Minimized;
            _actionController.RequestShowElement(item);
        }

        private void DeleteRemnants_Click(object sender, RoutedEventArgs e)
        {
            ApartmentExecutionReportItem item = GetReportItem(sender);
            if (item == null || _actionController == null)
                return;

            MessageBoxResult result = MessageBox.Show(
                this,
                "Удалить построенные элементы этой квартиры из отчёта?",
                "Удаление остатков",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
                return;

            _actionController.RequestDeleteRemnants(item);
        }

        private void Restore2DFamily_Click(object sender, RoutedEventArgs e)
        {
            ApartmentExecutionReportItem item = GetReportItem(sender);
            if (item == null || _actionController == null)
                return;

            _actionController.RequestRestore2DFamily(item);
        }

        private static ApartmentExecutionReportItem GetReportItem(object sender)
        {
            FrameworkElement source = sender as FrameworkElement;
            return source != null
                ? source.DataContext as ApartmentExecutionReportItem
                : null;
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
        public long NavigationElementId { get; set; }
        public ObservableCollection<long> NavigationElementIds { get; set; }
        public ObservableCollection<long> DeletableElementIds { get; set; }
        public Apartment2DRestoreInfo Restore2DInfo { get; set; }
        public string CustomHeaderText { get; set; }

        public long EffectiveNavigationElementId
        {
            get
            {
                return NavigationElementId > 0
                    ? NavigationElementId
                    : ApartmentId;
            }
        }

        public bool CanNavigateToElement
        {
            get { return GetNavigationCandidates().Count > 0; }
        }

        public bool CanDeleteRemnants
        {
            get { return DeletableElementIds != null && DeletableElementIds.Count > 0; }
        }

        public bool CanRestore2D
        {
            get { return Restore2DInfo != null && Restore2DInfo.CanRestore; }
        }

        public List<long> GetNavigationCandidates()
        {
            List<long> result = new List<long>();

            if (NavigationElementIds != null)
            {
                foreach (long value in NavigationElementIds)
                {
                    if (value > 0 && !result.Contains(value))
                        result.Add(value);
                }
            }

            if (NavigationElementId > 0 && !result.Contains(NavigationElementId))
                result.Add(NavigationElementId);

            if (ApartmentId > 0 && !result.Contains(ApartmentId))
                result.Add(ApartmentId);

            return result;
        }

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
            NavigationElementIds = new ObservableCollection<long>();
            DeletableElementIds = new ObservableCollection<long>();
        }
    }

    public class Apartment2DRestoreInfo
    {
        public long SymbolId { get; set; }
        public long ViewId { get; set; }
        public long LevelId { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public double Rotation { get; set; }

        public bool CanRestore
        {
            get { return SymbolId > 0 && ViewId > 0; }
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