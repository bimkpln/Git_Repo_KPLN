using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace KPLN_CoordiantorAI.Forms
{
    internal static class ModelChatMarkdownRenderer
    {
        public static void Render(
            RichTextBox richTextBox,
            string markdown,
            Brush foreground = null,
            Brush codeForeground = null,
            Brush tableHeaderBackground = null)
        {
            if (richTextBox == null)
                throw new ArgumentNullException(nameof(richTextBox));

            Brush effectiveForeground = foreground ?? Brushes.Black;
            Brush effectiveCodeForeground = codeForeground
                ?? new SolidColorBrush(Color.FromRgb(0, 112, 192));
            Brush effectiveHeaderBackground = tableHeaderBackground
                ?? new SolidColorBrush(Color.FromRgb(240, 240, 240));

            FlowDocument document = new FlowDocument
            {
                PagePadding = new Thickness(0),
                FontFamily = richTextBox.FontFamily,
                FontSize = richTextBox.FontSize,
                Foreground = effectiveForeground
            };
            richTextBox.Document = document;
            richTextBox.Foreground = effectiveForeground;
            if (string.IsNullOrEmpty(markdown))
                return;

            string[] lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            bool inTable = false;
            List<string[]> tableRows = new List<string[]>();
            int tableColumnCount = 0;

            foreach (string sourceLine in lines)
            {
                string line = sourceLine.Trim();
                if (line.StartsWith("|") && line.EndsWith("|"))
                {
                    if (IsTableSeparatorRow(line))
                        continue;

                    if (!inTable)
                    {
                        inTable = true;
                        tableRows.Clear();
                        tableColumnCount = 0;
                    }

                    string[] cells = line.Trim('|').Split('|')
                        .Select(cell => cell.Trim())
                        .ToArray();
                    tableColumnCount = Math.Max(tableColumnCount, cells.Length);
                    tableRows.Add(cells);
                    continue;
                }

                FlushTable(document, tableRows, ref inTable, tableColumnCount,
                    effectiveForeground, effectiveCodeForeground, effectiveHeaderBackground);

                if (string.IsNullOrWhiteSpace(line))
                    document.Blocks.Add(new Paragraph { Margin = new Thickness(0, 0, 0, 5) });
                else
                    AddRegularLine(document, line, effectiveCodeForeground);
            }

            FlushTable(document, tableRows, ref inTable, tableColumnCount,
                effectiveForeground, effectiveCodeForeground, effectiveHeaderBackground);
        }

        private static bool IsTableSeparatorRow(string line)
        {
            string[] cells = line.Trim('|').Split('|');
            return cells.Length > 0
                && cells.All(cell => Regex.IsMatch(cell, @"^\s*:?-+:?\s*$"));
        }

        private static void FlushTable(
            FlowDocument document,
            List<string[]> rows,
            ref bool inTable,
            int columnCount,
            Brush foreground,
            Brush codeForeground,
            Brush headerBackground)
        {
            if (!inTable || rows.Count == 0)
                return;

            Table table = new Table
            {
                CellSpacing = 0,
                BorderBrush = Brushes.Gray,
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 1, 0, 6),
                Foreground = foreground
            };
            for (int index = 0; index < columnCount; index++)
                table.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });

            TableRowGroup rowGroup = new TableRowGroup();
            table.RowGroups.Add(rowGroup);
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                TableRow row = new TableRow();
                if (rowIndex == 0)
                    row.Background = headerBackground;

                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    string cellText = columnIndex < rows[rowIndex].Length
                        ? rows[rowIndex][columnIndex]
                        : string.Empty;
                    Paragraph paragraph = new Paragraph { Margin = new Thickness(0) };
                    AddInlineMarkdown(paragraph, cellText, codeForeground);
                    TableCell cell = new TableCell(paragraph)
                    {
                        BorderBrush = Brushes.Gray,
                        BorderThickness = new Thickness(1),
                        Padding = new Thickness(5),
                        Foreground = foreground
                    };
                    if (rowIndex == 0)
                        cell.FontWeight = FontWeights.Bold;
                    row.Cells.Add(cell);
                }
                rowGroup.Rows.Add(row);
            }

            document.Blocks.Add(table);
            rows.Clear();
            inTable = false;
        }

        private static void AddRegularLine(FlowDocument document, string line, Brush codeForeground)
        {
            Paragraph paragraph = new Paragraph { Margin = new Thickness(0, 0, 0, 5) };
            if (TryAddHeading(document, paragraph, line, "### ", 14)
                || TryAddHeading(document, paragraph, line, "## ", 16)
                || TryAddHeading(document, paragraph, line, "# ", 18))
                return;

            if (line == "---" || line == "***")
            {
                paragraph.Inlines.Add(new InlineUIContainer(new Separator()));
                document.Blocks.Add(paragraph);
                return;
            }

            AddInlineMarkdown(paragraph, line, codeForeground);
            document.Blocks.Add(paragraph);
        }

        private static bool TryAddHeading(
            FlowDocument document,
            Paragraph paragraph,
            string line,
            string marker,
            double fontSize)
        {
            if (!line.StartsWith(marker, StringComparison.Ordinal))
                return false;

            paragraph.Inlines.Add(new Run(line.Substring(marker.Length))
            {
                FontWeight = FontWeights.Bold,
                FontSize = fontSize
            });
            document.Blocks.Add(paragraph);
            return true;
        }

        private static void AddInlineMarkdown(Paragraph paragraph, string text, Brush codeForeground)
        {
            if (string.IsNullOrEmpty(text))
                return;

            string codeMarker = ((char)96).ToString();
            int position = 0;
            while (position < text.Length)
            {
                int previousPosition = position;
                int boldStart = text.IndexOf("**", position, StringComparison.Ordinal);
                int italicStart = text.IndexOf("*", position, StringComparison.Ordinal);
                int codeStart = text.IndexOf(codeMarker, position, StringComparison.Ordinal);
                int nextMarker = -1;
                string markerType = null;

                if (boldStart >= 0)
                {
                    nextMarker = boldStart;
                    markerType = "bold";
                }
                if (italicStart >= 0
                    && (nextMarker < 0 || italicStart < nextMarker)
                    && italicStart + 1 < text.Length
                    && text[italicStart + 1] != '*')
                {
                    nextMarker = italicStart;
                    markerType = "italic";
                }
                if (codeStart >= 0 && (nextMarker < 0 || codeStart < nextMarker))
                {
                    nextMarker = codeStart;
                    markerType = "code";
                }

                if (nextMarker < 0)
                {
                    paragraph.Inlines.Add(new Run(text.Substring(position)));
                    break;
                }
                if (nextMarker > position)
                    paragraph.Inlines.Add(new Run(text.Substring(position, nextMarker - position)));

                position = nextMarker;
                if (markerType == "bold")
                {
                    int end = text.IndexOf("**", position + 2, StringComparison.Ordinal);
                    if (end >= 0)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position + 2, end - position - 2))
                        {
                            FontWeight = FontWeights.Bold
                        });
                        position = end + 2;
                    }
                    else
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position)));
                        position = text.Length;
                    }
                }
                else if (markerType == "italic")
                {
                    int end = text.IndexOf("*", position + 1, StringComparison.Ordinal);
                    if (end >= 0)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position + 1, end - position - 1))
                        {
                            FontStyle = FontStyles.Italic
                        });
                        position = end + 1;
                    }
                    else
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position)));
                        position = text.Length;
                    }
                }
                else
                {
                    int end = text.IndexOf(codeMarker, position + 1, StringComparison.Ordinal);
                    if (end >= 0)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position + 1, end - position - 1))
                        {
                            FontFamily = new FontFamily("Consolas"),
                            Foreground = codeForeground
                        });
                        position = end + 1;
                    }
                    else
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(position, 1)));
                        position++;
                    }
                }

                if (position == previousPosition)
                {
                    paragraph.Inlines.Add(new Run(text.Substring(position, 1)));
                    position++;
                }
            }
        }
    }
}
