using System;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using KPLN_CoordiantorAI.Forms;

internal static class ModelChatMarkdownRendererTests
{
    [STAThread]
    private static int Main()
    {
        try
        {
            RichTextBox richTextBox = new RichTextBox
            {
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13
            };
            string codeMarker = ((char)96).ToString();
            string markdown =
                "## Heading\nNormal **bold** and *italic* and "
                + codeMarker + "code" + codeMarker
                + "\n\n| A | B |\n| --- | :---: |\n| 1 | 2 |";

            ModelChatMarkdownRenderer.Render(
                richTextBox,
                markdown,
                Brushes.White,
                Brushes.LightBlue,
                Brushes.DarkSlateGray);

            Paragraph heading = richTextBox.Document.Blocks.FirstBlock as Paragraph;
            Run headingRun = heading == null ? null : heading.Inlines.FirstInline as Run;
            Check(headingRun != null && headingRun.Text == "Heading", "Heading marker was not removed.");
            Check(headingRun.FontWeight == System.Windows.FontWeights.Bold, "Heading is not bold.");
            Check(Math.Abs(headingRun.FontSize - 16) < 0.01, "Wrong heading size.");

            Paragraph formatted = heading.NextBlock as Paragraph;
            Run[] runs = formatted.Inlines.OfType<Run>().ToArray();
            Check(runs.Any(run => run.Text == "bold" && run.FontWeight == System.Windows.FontWeights.Bold),
                "Bold Markdown was not rendered.");
            Check(runs.Any(run => run.Text == "italic" && run.FontStyle == System.Windows.FontStyles.Italic),
                "Italic Markdown was not rendered.");
            Check(runs.Any(run => run.Text == "code" && run.FontFamily.Source == "Consolas"),
                "Inline code was not rendered.");

            Table table = richTextBox.Document.Blocks.OfType<Table>().Single();
            Check(table.RowGroups.Count == 1 && table.RowGroups[0].Rows.Count == 2,
                "Markdown table separator became a data row.");
            Check(table.RowGroups[0].Rows[0].Cells.Count == 2, "Wrong table column count.");
            Check(richTextBox.Document.Foreground == Brushes.White, "Dark-theme foreground was not applied.");

            RichTextBox empty = new RichTextBox();
            ModelChatMarkdownRenderer.Render(empty, null);
            Check(empty.Document.Blocks.Count == 0, "Empty Markdown created content.");

            Console.WriteLine("All Markdown renderer tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
