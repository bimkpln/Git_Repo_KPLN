using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_ModelChecker_User.Forms
{
    internal enum SpecSumReportItemKind
    {
        Warning,
        Title,
        ScheduleBlock,
        TotalLine
    }

    internal enum SpecSumReportSection
    {
        None,
        Main,
        GeneralData
    }

    internal enum SpecSumTotalRole
    {
        None,
        MainTotal,
        GeneralDataTotal,
        DifferenceValue,
        DifferencePercent
    }

    internal sealed class SpecSumReportItem
    {
        internal SpecSumReportItemKind Kind { get; set; }
        internal SpecSumReportSection Section { get; set; }
        internal SpecSumTotalRole TotalRole { get; set; }
        internal string Title { get; set; }
        internal string Message { get; set; }
        internal string SheetNumber { get; set; }
        internal string SheetName { get; set; }
        internal string ScheduleName { get; set; }
        internal string Label { get; set; }
        internal string ValueText { get; set; }
        internal bool Muted { get; set; }
        internal bool Accent { get; set; }
        internal bool HasNumericValue { get; set; }
        internal double BaseValue { get; set; }
        internal int Multiplier { get; set; }
        internal bool IsExcluded { get; set; }
        internal int FilteredOutRows { get; set; }
        internal double FilteredOutValue { get; set; }
    }

    public partial class SpecSumReportWindow : Window
    {
        private const string ExcludedDesignationParamName = "О_Обозначение";

        private readonly string _reportText;
        private readonly List<SpecSumReportItem> _reportItems;
        private readonly Dictionary<SpecSumReportItem, Border> _scheduleBorders =
            new Dictionary<SpecSumReportItem, Border>();
        private readonly Dictionary<SpecSumReportItem, TextBlock> _scheduleValueTextBlocks =
            new Dictionary<SpecSumReportItem, TextBlock>();
        private readonly Dictionary<SpecSumReportItem, TextBlock> _scheduleStatusTextBlocks =
            new Dictionary<SpecSumReportItem, TextBlock>();
        private readonly Dictionary<SpecSumReportItem, Button> _scheduleExcludeButtons =
            new Dictionary<SpecSumReportItem, Button>();
        private readonly Dictionary<SpecSumTotalRole, TextBlock> _totalValueTextBlocks =
            new Dictionary<SpecSumTotalRole, TextBlock>();

        private bool _isUpdatingMultiplierText;

        internal SpecSumReportWindow(
            string title,
            IEnumerable<SpecSumReportItem> reportItems,
            string reportText)
        {
            InitializeComponent();

            Title = string.IsNullOrEmpty(title) ? "Сумма спецификаций" : title;
            HeaderTextBlock.Text = Title;

            _reportText = reportText ?? string.Empty;
            _reportItems = reportItems == null
                ? new List<SpecSumReportItem>()
                : reportItems.ToList();

            NormalizeReportItems();
            ReportViewer.Document = BuildDocument();
            RecalculateTotals();
        }

        private void NormalizeReportItems()
        {
            foreach (SpecSumReportItem item in _reportItems)
            {
                if (item.Kind != SpecSumReportItemKind.ScheduleBlock)
                {
                    continue;
                }

                if (!item.HasNumericValue)
                {
                    item.Multiplier = 1;
                    continue;
                }

                if (item.Multiplier < 0)
                {
                    item.Multiplier = 1;
                }
            }
        }

        private FlowDocument BuildDocument()
        {
            _scheduleBorders.Clear();
            _scheduleValueTextBlocks.Clear();
            _scheduleStatusTextBlocks.Clear();
            _scheduleExcludeButtons.Clear();
            _totalValueTextBlocks.Clear();

            FlowDocument document = new FlowDocument
            {
                PagePadding = new Thickness(18),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                Background = Brushes.White
            };

            if (_reportItems.Count == 0)
            {
                document.Blocks.Add(new Paragraph(new Run(_reportText)));
                return document;
            }

            foreach (SpecSumReportItem item in _reportItems)
            {
                switch (item.Kind)
                {
                    case SpecSumReportItemKind.Warning:
                        AddWarning(document, item.Message);
                        break;
                    case SpecSumReportItemKind.Title:
                        AddTitle(document, item.Title);
                        break;
                    case SpecSumReportItemKind.ScheduleBlock:
                        AddScheduleBlock(document, item);
                        break;
                    case SpecSumReportItemKind.TotalLine:
                        AddTotalLine(document, item);
                        break;
                }
            }

            return document;
        }

        private static void AddWarning(FlowDocument document, string message)
        {
            Border border = new Border
            {
                Background = BrushFromHex("#fff8e1"),
                BorderBrush = BrushFromHex("#ffe082"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 0, 0, 14)
            };

            border.Child = new TextBlock
            {
                Text = "ВНИМАНИЕ: " + (message ?? string.Empty),
                Foreground = BrushFromHex("#7d4e00"),
                FontSize = 14,
                TextWrapping = TextWrapping.Wrap
            };

            document.Blocks.Add(new BlockUIContainer(border));
        }

        private static void AddTitle(FlowDocument document, string title)
        {
            if (document.Blocks.Count > 0)
            {
                document.Blocks.Add(new Paragraph { Margin = new Thickness(0, 4, 0, 0) });
            }

            Paragraph paragraph = new Paragraph(new Run(title ?? string.Empty))
            {
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                Foreground = BrushFromHex("#1f2d3d"),
                Margin = new Thickness(0, 10, 0, 6)
            };

            document.Blocks.Add(paragraph);
            document.Blocks.Add(new BlockUIContainer(new Border
            {
                Height = 2,
                Background = BrushFromHex("#d8e2ee"),
                Margin = new Thickness(0, 0, 0, 16)
            }));
        }

        private void AddScheduleBlock(FlowDocument document, SpecSumReportItem item)
        {
            StackPanel panel = new StackPanel();

            panel.Children.Add(new TextBlock
            {
                Text = string.Format("{0} | {1}", item.SheetNumber, item.SheetName),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = BrushFromHex("#1f2d3d"),
                TextWrapping = TextWrapping.Wrap
            });

            panel.Children.Add(new Border
            {
                Height = 1,
                Background = BrushFromHex("#dde6ef"),
                Margin = new Thickness(0, 8, 0, 8)
            });

            panel.Children.Add(new TextBlock
            {
                Text = item.ScheduleName ?? string.Empty,
                FontWeight = FontWeights.SemiBold,
                Foreground = BrushFromHex("#243447"),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            TextBlock valueTextBlock = new TextBlock
            {
                Text = item.ValueText ?? string.Empty,
                Foreground = item.Muted ? BrushFromHex("#6b7280") : BrushFromHex("#0b57d0"),
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            _scheduleValueTextBlocks[item] = valueTextBlock;

            if (item.HasNumericValue)
            {
                Grid valueGrid = new Grid();
                valueGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                valueGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                Grid.SetColumn(valueTextBlock, 0);
                valueGrid.Children.Add(valueTextBlock);

                StackPanel controls = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(16, 0, 0, 0)
                };

                controls.Children.Add(new TextBlock
                {
                    Text = "x",
                    Margin = new Thickness(0, 0, 6, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Foreground = BrushFromHex("#4b5563")
                });

                TextBox multiplierTextBox = new TextBox
                {
                    Text = item.Multiplier.ToString(CultureInfo.InvariantCulture),
                    Tag = item,
                    Width = 56,
                    Height = 26,
                    MaxLength = 6,
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                };
                multiplierTextBox.PreviewTextInput += MultiplierTextBox_PreviewTextInput;
                multiplierTextBox.TextChanged += MultiplierTextBox_TextChanged;
                DataObject.AddPastingHandler(multiplierTextBox, MultiplierTextBox_Pasting);
                controls.Children.Add(multiplierTextBox);

                Button excludeButton = new Button
                {
                    Tag = item,
                    MinWidth = 92,
                    Height = 26,
                    Padding = new Thickness(8, 0, 8, 0)
                };
                excludeButton.Click += ExcludeButton_Click;
                controls.Children.Add(excludeButton);
                _scheduleExcludeButtons[item] = excludeButton;

                Grid.SetColumn(controls, 1);
                valueGrid.Children.Add(controls);
                panel.Children.Add(valueGrid);

                TextBlock statusTextBlock = new TextBlock
                {
                    Margin = new Thickness(0, 8, 0, 0),
                    FontSize = 12,
                    TextWrapping = TextWrapping.Wrap
                };
                panel.Children.Add(statusTextBlock);
                _scheduleStatusTextBlocks[item] = statusTextBlock;
            }
            else
            {
                panel.Children.Add(valueTextBlock);
            }

            if (item.FilteredOutRows > 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = string.Format(
                        CultureInfo.InvariantCulture,
                        "Исключено по '{0}': {1} строк, {2}.",
                        ExcludedDesignationParamName,
                        item.FilteredOutRows,
                        FormatNumber(item.FilteredOutValue)),
                    Margin = new Thickness(0, 8, 0, 0),
                    Foreground = BrushFromHex("#7d4e00"),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            Border border = new Border
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12, 10, 12, 12),
                Margin = new Thickness(0, 0, 0, 18),
                Child = panel
            };
            _scheduleBorders[item] = border;

            document.Blocks.Add(new BlockUIContainer(border));
            UpdateScheduleBlockVisual(item);
        }

        private void AddTotalLine(FlowDocument document, SpecSumReportItem item)
        {
            Brush borderBrush = item.Accent ? BrushFromHex("#0b57d0") : BrushFromHex("#8a94a6");
            Brush valueBrush = item.Accent ? BrushFromHex("#0b57d0") : BrushFromHex("#4b5563");

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(4) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Border leftBorder = new Border { Background = borderBrush };
            Grid.SetColumn(leftBorder, 0);
            grid.Children.Add(leftBorder);

            StackPanel line = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(12, 10, 12, 10)
            };

            line.Children.Add(new TextBlock
            {
                Text = item.Label ?? string.Empty,
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = BrushFromHex("#1f2d3d"),
                VerticalAlignment = VerticalAlignment.Center
            });

            TextBlock valueTextBlock = new TextBlock
            {
                Text = item.ValueText ?? string.Empty,
                Margin = new Thickness(18, 0, 0, 0),
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = valueBrush,
                VerticalAlignment = VerticalAlignment.Center
            };
            line.Children.Add(valueTextBlock);

            if (item.TotalRole != SpecSumTotalRole.None)
            {
                _totalValueTextBlocks[item.TotalRole] = valueTextBlock;
            }

            Grid.SetColumn(line, 1);
            grid.Children.Add(line);

            Border border = new Border
            {
                Background = BrushFromHex("#fbfcfe"),
                Margin = new Thickness(0, 14, 0, 10),
                Child = grid
            };

            document.Blocks.Add(new BlockUIContainer(border));
        }

        private void ExcludeButton_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            SpecSumReportItem item = button == null ? null : button.Tag as SpecSumReportItem;
            if (item == null)
            {
                return;
            }

            item.IsExcluded = !item.IsExcluded;
            RecalculateTotals();
        }

        private void MultiplierTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsDigits(e.Text);
        }

        private void MultiplierTextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(DataFormats.Text))
            {
                e.CancelCommand();
                return;
            }

            string text = e.DataObject.GetData(DataFormats.Text) as string;
            if (!IsDigits(text))
            {
                e.CancelCommand();
            }
        }

        private void MultiplierTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingMultiplierText)
            {
                return;
            }

            TextBox textBox = sender as TextBox;
            SpecSumReportItem item = textBox == null ? null : textBox.Tag as SpecSumReportItem;
            if (item == null)
            {
                return;
            }

            string text = textBox.Text ?? string.Empty;
            if (text.Length == 0)
            {
                item.Multiplier = 0;
                RecalculateTotals();
                return;
            }

            int multiplier;
            if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out multiplier)
                && multiplier >= 0)
            {
                item.Multiplier = multiplier;
                RecalculateTotals();
                return;
            }

            string cleaned = new string(text.Where(char.IsDigit).ToArray());
            if (cleaned.Length == 0)
            {
                cleaned = "0";
            }

            _isUpdatingMultiplierText = true;
            textBox.Text = cleaned;
            textBox.CaretIndex = textBox.Text.Length;
            _isUpdatingMultiplierText = false;
            MultiplierTextBox_TextChanged(sender, e);
        }

        private static bool IsDigits(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            return text.All(char.IsDigit);
        }

        private void RecalculateTotals()
        {
            foreach (SpecSumReportItem item in _reportItems)
            {
                if (item.Kind == SpecSumReportItemKind.ScheduleBlock)
                {
                    UpdateScheduleBlockVisual(item);
                }
            }

            UpdateTotalValue(SpecSumTotalRole.MainTotal, FormatNumber(GetSectionTotal(SpecSumReportSection.Main)));
            UpdateTotalValue(SpecSumTotalRole.GeneralDataTotal, FormatNumber(GetSectionTotal(SpecSumReportSection.GeneralData)));

            double mainTotal = GetSectionTotal(SpecSumReportSection.Main);
            double generalDataTotal = GetSectionTotal(SpecSumReportSection.GeneralData);
            UpdateTotalValue(SpecSumTotalRole.DifferenceValue, FormatNumber(Math.Abs(mainTotal - generalDataTotal)));

            string differencePercentText = generalDataTotal > 0.0
                ? FormatPercent(Math.Abs(mainTotal - generalDataTotal) / generalDataTotal * 100.0)
                : "нет данных";
            UpdateTotalValue(SpecSumTotalRole.DifferencePercent, differencePercentText);
        }

        private void UpdateTotalValue(SpecSumTotalRole role, string valueText)
        {
            TextBlock textBlock;
            if (_totalValueTextBlocks.TryGetValue(role, out textBlock))
            {
                textBlock.Text = valueText;
            }

            foreach (SpecSumReportItem item in _reportItems)
            {
                if (item.Kind == SpecSumReportItemKind.TotalLine && item.TotalRole == role)
                {
                    item.ValueText = valueText;
                }
            }
        }

        private void UpdateScheduleBlockVisual(SpecSumReportItem item)
        {
            TextBlock valueTextBlock;
            if (_scheduleValueTextBlocks.TryGetValue(item, out valueTextBlock))
            {
                valueTextBlock.Text = GetScheduleDisplayText(item);
                valueTextBlock.Foreground = item.IsExcluded || item.Muted
                    ? BrushFromHex("#6b7280")
                    : BrushFromHex("#0b57d0");
            }

            TextBlock statusTextBlock;
            if (_scheduleStatusTextBlocks.TryGetValue(item, out statusTextBlock))
            {
                if (item.IsExcluded)
                {
                    statusTextBlock.Text = "Не учитывается в итогах.";
                    statusTextBlock.Foreground = BrushFromHex("#6b7280");
                }
                else if (item.Multiplier != 1)
                {
                    statusTextBlock.Text = string.Format(
                        CultureInfo.InvariantCulture,
                        "Учитывается с множителем x{0}.",
                        item.Multiplier);
                    statusTextBlock.Foreground = BrushFromHex("#4b5563");
                }
                else
                {
                    statusTextBlock.Text = "Учитывается в итогах.";
                    statusTextBlock.Foreground = BrushFromHex("#4b5563");
                }
            }

            Button excludeButton;
            if (_scheduleExcludeButtons.TryGetValue(item, out excludeButton))
            {
                excludeButton.Content = item.IsExcluded ? "Вернуть" : "Исключить";
            }

            Border border;
            if (_scheduleBorders.TryGetValue(item, out border))
            {
                if (item.IsExcluded)
                {
                    border.Background = BrushFromHex("#f3f4f6");
                    border.BorderBrush = BrushFromHex("#d1d5db");
                    border.Opacity = 0.68;
                }
                else
                {
                    border.Background = BrushFromHex("#fbfcfe");
                    border.BorderBrush = BrushFromHex("#e3ebf3");
                    border.Opacity = 1.0;
                }
            }
        }

        private double GetSectionTotal(SpecSumReportSection section)
        {
            double total = 0.0;

            foreach (SpecSumReportItem item in _reportItems)
            {
                if (item.Kind == SpecSumReportItemKind.ScheduleBlock
                    && item.Section == section
                    && item.HasNumericValue
                    && !item.IsExcluded)
                {
                    total += item.BaseValue * item.Multiplier;
                }
            }

            return total;
        }

        private static string GetScheduleDisplayText(SpecSumReportItem item)
        {
            if (!item.HasNumericValue)
            {
                return item.ValueText ?? string.Empty;
            }

            double multipliedValue = item.BaseValue * item.Multiplier;
            if (item.IsExcluded)
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    "Не учитывается: {0} ({1} x {2})",
                    FormatNumber(multipliedValue),
                    FormatNumber(item.BaseValue),
                    item.Multiplier);
            }

            if (item.Multiplier == 1)
            {
                return FormatNumber(item.BaseValue);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} ({1} x {2})",
                FormatNumber(multipliedValue),
                FormatNumber(item.BaseValue),
                item.Multiplier);
        }

        private static string GetScheduleExportText(SpecSumReportItem item)
        {
            if (!item.HasNumericValue)
            {
                return item.ValueText ?? string.Empty;
            }

            double multipliedValue = item.BaseValue * item.Multiplier;
            if (item.IsExcluded)
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    "НЕ УЧИТЫВАЕТСЯ: {0} ({1} x {2})",
                    FormatNumber(multipliedValue),
                    FormatNumber(item.BaseValue),
                    item.Multiplier);
            }

            if (item.Multiplier == 1)
            {
                return FormatNumber(item.BaseValue);
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} ({1} x {2})",
                FormatNumber(multipliedValue),
                FormatNumber(item.BaseValue),
                item.Multiplier);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Сохранить отчет",
                Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*",
                FileName = string.Format("SpecSum_{0:yyyyMMdd_HHmmss}.txt", DateTime.Now),
                AddExtension = true,
                DefaultExt = ".txt",
                OverwritePrompt = true
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
            {
                return;
            }

            File.WriteAllText(dialog.FileName, BuildTextReport(), Encoding.UTF8);

            MessageBox.Show(
                this,
                "Отчет сохранен.",
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void SaveHtmlButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                Title = "Сохранить отчет HTML",
                Filter = "HTML-файл (*.html)|*.html|Все файлы (*.*)|*.*",
                FileName = string.Format("SpecSum_{0:yyyyMMdd_HHmmss}.html", DateTime.Now),
                AddExtension = true,
                DefaultExt = ".html",
                OverwritePrompt = true
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
            {
                return;
            }

            File.WriteAllText(dialog.FileName, BuildHtmlReport(), Encoding.UTF8);

            MessageBox.Show(
                this,
                "HTML-отчет сохранен.",
                Title,
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private string BuildTextReport()
        {
            RecalculateTotals();

            if (_reportItems.Count == 0)
            {
                return _reportText;
            }

            StringBuilder output = new StringBuilder();

            foreach (SpecSumReportItem item in _reportItems)
            {
                switch (item.Kind)
                {
                    case SpecSumReportItemKind.Warning:
                        output.AppendLine("ВНИМАНИЕ: " + (item.Message ?? string.Empty));
                        output.AppendLine();
                        break;
                    case SpecSumReportItemKind.Title:
                        AppendTextTitle(output, item.Title);
                        break;
                    case SpecSumReportItemKind.ScheduleBlock:
                        AppendTextScheduleBlock(output, item);
                        break;
                    case SpecSumReportItemKind.TotalLine:
                        output.AppendLine((item.Label ?? string.Empty) + " " + (item.ValueText ?? string.Empty));
                        output.AppendLine();
                        break;
                }
            }

            AppendTextExcludedSummary(output);
            return output.ToString();
        }

        private static void AppendTextTitle(StringBuilder output, string title)
        {
            title = title ?? string.Empty;
            if (output.Length > 0)
            {
                output.AppendLine();
            }

            output.AppendLine(title);
            output.AppendLine(new string('=', title.Length));
            output.AppendLine();
        }

        private static void AppendTextScheduleBlock(StringBuilder output, SpecSumReportItem item)
        {
            output.AppendLine(string.Format("{0} | {1}", item.SheetNumber, item.SheetName));
            output.AppendLine("  " + (item.ScheduleName ?? string.Empty));
            output.AppendLine("  " + GetScheduleExportText(item));

            if (item.FilteredOutRows > 0)
            {
                output.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    "  Исключено по '{0}': {1} строк, {2}.",
                    ExcludedDesignationParamName,
                    item.FilteredOutRows,
                    FormatNumber(item.FilteredOutValue)));
            }

            output.AppendLine();
        }

        private void AppendTextExcludedSummary(StringBuilder output)
        {
            List<SpecSumReportItem> excludedItems = _reportItems
                .Where(x => x.Kind == SpecSumReportItemKind.ScheduleBlock && x.HasNumericValue && x.IsExcluded)
                .ToList();

            if (excludedItems.Count > 0)
            {
                output.AppendLine();
                output.AppendLine("Не учитывается в расчетах:");
                foreach (SpecSumReportItem item in excludedItems)
                {
                    output.AppendLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "- {0} | {1} | {2}: {3}",
                        item.SheetNumber,
                        item.SheetName,
                        item.ScheduleName,
                        GetScheduleExportText(item)));
                }
            }

            List<SpecSumReportItem> filteredItems = _reportItems
                .Where(x => x.Kind == SpecSumReportItemKind.ScheduleBlock && x.FilteredOutRows > 0)
                .ToList();

            if (filteredItems.Count > 0)
            {
                output.AppendLine();
                output.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    "Автоматически исключено по '{0}':",
                    ExcludedDesignationParamName));

                foreach (SpecSumReportItem item in filteredItems)
                {
                    output.AppendLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "- {0} | {1} | {2}: {3} строк, {4}",
                        item.SheetNumber,
                        item.SheetName,
                        item.ScheduleName,
                        item.FilteredOutRows,
                        FormatNumber(item.FilteredOutValue)));
                }
            }
        }

        private string BuildHtmlReport()
        {
            RecalculateTotals();

            StringBuilder html = new StringBuilder();

            html.Append("<!doctype html><html><head><meta charset=\"utf-8\"><title>");
            html.Append(HtmlEncode(Title));
            html.Append("</title></head><body style=\"margin:24px; background:#ffffff;\">");

            if (_reportItems.Count == 0)
            {
                html.Append("<pre style=\"font-family:Consolas, monospace; font-size:13px; color:#243447; white-space:pre-wrap;\">");
                html.Append(HtmlEncode(_reportText));
                html.Append("</pre>");
            }
            else
            {
                foreach (SpecSumReportItem item in _reportItems)
                {
                    switch (item.Kind)
                    {
                        case SpecSumReportItemKind.Warning:
                            AppendWarningHtml(html, item.Message);
                            break;
                        case SpecSumReportItemKind.Title:
                            AppendTitleHtml(html, item.Title);
                            break;
                        case SpecSumReportItemKind.ScheduleBlock:
                            AppendScheduleBlockHtml(html, item);
                            break;
                        case SpecSumReportItemKind.TotalLine:
                            AppendTotalLineHtml(html, item.Label, item.ValueText, item.Accent);
                            break;
                    }
                }

                AppendHtmlExcludedSummary(html);
            }

            html.Append("</body></html>");
            return html.ToString();
        }

        private static void AppendWarningHtml(StringBuilder html, string message)
        {
            html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:#7d4e00; background:#fff8e1; border:1px solid #ffe082; border-radius:4px; padding:8px 12px; margin:0 0 14px 0;\">");
            html.Append("ВНИМАНИЕ: ");
            html.Append(HtmlEncode(message));
            html.Append("</div>");
        }

        private static void AppendTitleHtml(StringBuilder html, string title)
        {
            html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; font-size:22px; font-weight:700; color:#1f2d3d; margin:10px 0 16px 0; padding-bottom:6px; border-bottom:2px solid #d8e2ee;\">");
            html.Append(HtmlEncode(title));
            html.Append("</div>");
        }

        private static void AppendScheduleBlockHtml(StringBuilder html, SpecSumReportItem item)
        {
            string blockStyle = item.IsExcluded
                ? "background:#f3f4f6; border:1px solid #d1d5db; color:#6b7280; opacity:.78;"
                : "background:#fbfcfe; border:1px solid #e3ebf3; color:#243447;";
            string valueStyle = item.IsExcluded || item.Muted
                ? "color:#6b7280; font-weight:700; user-select:text;"
                : "color:#0b57d0; font-weight:700; user-select:text;";

            html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; margin:0 0 18px 0; padding:10px 12px 12px 12px; border-radius:4px; font-size:17px; ");
            html.Append(blockStyle);
            html.Append("\">");
            html.Append("<div style=\"font-size:18px; font-weight:700; color:#1f2d3d; margin:0 0 12px 0; padding:0 0 8px 0; border-bottom:1px solid #dde6ef;\">");
            html.Append(HtmlEncode(string.Format("{0} | {1}", item.SheetNumber, item.SheetName)));
            html.Append("</div>");
            html.Append("<div style=\"font-weight:600; line-height:1.35; margin:0 0 8px 0;\">");
            html.Append(HtmlEncode(item.ScheduleName));
            html.Append("</div>");
            html.Append("<div style=\"");
            html.Append(valueStyle);
            html.Append("\">");
            html.Append(HtmlEncode(GetScheduleExportText(item)));
            html.Append("</div>");

            if (item.FilteredOutRows > 0)
            {
                html.Append("<div style=\"margin-top:8px; color:#7d4e00; font-size:14px;\">");
                html.Append(HtmlEncode(string.Format(
                    CultureInfo.InvariantCulture,
                    "Исключено по '{0}': {1} строк, {2}.",
                    ExcludedDesignationParamName,
                    item.FilteredOutRows,
                    FormatNumber(item.FilteredOutValue))));
                html.Append("</div>");
            }

            html.Append("</div>");
        }

        private static void AppendTotalLineHtml(
            StringBuilder html,
            string label,
            string valueText,
            bool accent)
        {
            string borderColor = accent ? "#0b57d0" : "#8a94a6";
            string valueColor = accent ? "#0b57d0" : "#4b5563";

            html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; font-size:20px; font-weight:700; color:#1f2d3d; margin:14px 0 10px 0; padding:10px 12px; background:#fbfcfe; border-left:4px solid ");
            html.Append(borderColor);
            html.Append("; white-space:nowrap;\">");
            html.Append("<span>");
            html.Append(HtmlEncode(label));
            html.Append("</span><span style=\"display:inline-block; min-width:18px;\"></span><span style=\"color:");
            html.Append(valueColor);
            html.Append("; font-weight:700; user-select:text;\">");
            html.Append(HtmlEncode(valueText));
            html.Append("</span></div>");
        }

        private void AppendHtmlExcludedSummary(StringBuilder html)
        {
            List<SpecSumReportItem> excludedItems = _reportItems
                .Where(x => x.Kind == SpecSumReportItemKind.ScheduleBlock && x.HasNumericValue && x.IsExcluded)
                .ToList();

            if (excludedItems.Count > 0)
            {
                AppendTitleHtml(html, "Не учитывается в расчетах");
                foreach (SpecSumReportItem item in excludedItems)
                {
                    AppendScheduleBlockHtml(html, item);
                }
            }

            List<SpecSumReportItem> filteredItems = _reportItems
                .Where(x => x.Kind == SpecSumReportItemKind.ScheduleBlock && x.FilteredOutRows > 0)
                .ToList();

            if (filteredItems.Count > 0)
            {
                AppendTitleHtml(
                    html,
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "Автоматически исключено по '{0}'",
                        ExcludedDesignationParamName));

                foreach (SpecSumReportItem item in filteredItems)
                {
                    html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; margin:0 0 8px 0; padding:8px 12px; background:#fff8e1; border:1px solid #ffe082; border-radius:4px; color:#7d4e00;\">");
                    html.Append(HtmlEncode(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} | {1} | {2}: {3} строк, {4}",
                        item.SheetNumber,
                        item.SheetName,
                        item.ScheduleName,
                        item.FilteredOutRows,
                        FormatNumber(item.FilteredOutValue))));
                    html.Append("</div>");
                }
            }
        }

        private static Brush BrushFromHex(string hex)
        {
            return (Brush)new BrushConverter().ConvertFromString(hex);
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", ",") + " кг";
        }

        private static string FormatPercent(double value)
        {
            return value.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", ",") + "%";
        }

        private static string HtmlEncode(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
        }
    }
}