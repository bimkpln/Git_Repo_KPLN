using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace KPLN_Tools.Forms
{
    internal enum SpecSumReportItemKind
    {
        Warning,
        Title,
        ScheduleBlock,
        TotalLine
    }

    internal sealed class SpecSumReportItem
    {
        internal SpecSumReportItemKind Kind { get; set; }
        internal string Title { get; set; }
        internal string Message { get; set; }
        internal string SheetNumber { get; set; }
        internal string SheetName { get; set; }
        internal string ScheduleName { get; set; }
        internal string Label { get; set; }
        internal string ValueText { get; set; }
        internal bool Muted { get; set; }
        internal bool Accent { get; set; }
    }

    public partial class SpecSumReportWindow : Window
    {
        private readonly string _reportText;
        private readonly List<SpecSumReportItem> _reportItems;

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

            ReportViewer.Document = BuildDocument();
        }

        private FlowDocument BuildDocument()
        {
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
                        AddTotalLine(document, item.Label, item.ValueText, item.Accent);
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

        private static void AddScheduleBlock(FlowDocument document, SpecSumReportItem item)
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

            panel.Children.Add(new TextBlock
            {
                Text = item.ValueText ?? string.Empty,
                Foreground = item.Muted ? BrushFromHex("#6b7280") : BrushFromHex("#0b57d0"),
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap
            });

            Border border = new Border
            {
                Background = BrushFromHex("#fbfcfe"),
                BorderBrush = BrushFromHex("#e3ebf3"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12, 10, 12, 12),
                Margin = new Thickness(0, 0, 0, 18),
                Child = panel
            };

            document.Blocks.Add(new BlockUIContainer(border));
        }

        private static void AddTotalLine(
            FlowDocument document,
            string label,
            string valueText,
            bool accent)
        {
            Brush borderBrush = accent ? BrushFromHex("#0b57d0") : BrushFromHex("#8a94a6");
            Brush valueBrush = accent ? BrushFromHex("#0b57d0") : BrushFromHex("#4b5563");

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
                Text = label ?? string.Empty,
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = BrushFromHex("#1f2d3d"),
                VerticalAlignment = VerticalAlignment.Center
            });

            line.Children.Add(new TextBlock
            {
                Text = valueText ?? string.Empty,
                Margin = new Thickness(18, 0, 0, 0),
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = valueBrush,
                VerticalAlignment = VerticalAlignment.Center
            });

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

        private static Brush BrushFromHex(string hex)
        {
            return (Brush)new BrushConverter().ConvertFromString(hex);
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

            File.WriteAllText(dialog.FileName, _reportText, Encoding.UTF8);

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

        private string BuildHtmlReport()
        {
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
            string valueStyle = item.Muted
                ? "color:#6b7280; font-weight:700; user-select:text;"
                : "color:#0b57d0; font-weight:700; user-select:text;";

            html.Append("<div style=\"font-family:Segoe UI, Arial, sans-serif; margin:0 0 18px 0; padding:10px 12px 12px 12px; background:#fbfcfe; border:1px solid #e3ebf3; border-radius:4px; font-size:17px; color:#243447;\">");
            html.Append("<div style=\"font-size:18px; font-weight:700; color:#1f2d3d; margin:0 0 12px 0; padding:0 0 8px 0; border-bottom:1px solid #dde6ef;\">");
            html.Append(HtmlEncode(string.Format("{0} | {1}", item.SheetNumber, item.SheetName)));
            html.Append("</div>");
            html.Append("<div style=\"font-weight:600; line-height:1.35; margin:0 0 8px 0;\">");
            html.Append(HtmlEncode(item.ScheduleName));
            html.Append("</div>");
            html.Append("<div style=\"");
            html.Append(valueStyle);
            html.Append("\">");
            html.Append(HtmlEncode(item.ValueText));
            html.Append("</div></div>");
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

        private static string HtmlEncode(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
        }
    }
}

