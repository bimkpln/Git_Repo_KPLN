using KPLN_Publication;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Path = System.IO.Path;

internal static class Program
{
    private static int _checks;
    private static void Assert(bool condition, string description)
    {
        if (!condition) throw new Exception(description);
        _checks++;
    }

    private static SheetPrintBounds Mm(double x1, double y1, double x2, double y2)
    {
        return new SheetPrintBounds(x1 / 304.8, y1 / 304.8, x2 / 304.8, y2 / 304.8);
    }

    private static void MakePdf(string path, string text)
    {
        using (var stream = File.Create(path))
        using (var doc = new Document())
        using (var writer = PdfWriter.GetInstance(doc, stream))
        {
            doc.Open();
            doc.Add(new Paragraph(text));
            doc.Close();
        }
    }

    private static int Main()
    {
        try
        {
            // Регрессия на координатах листа насосной из модели пользователя.
            var frame = Mm(-1261, 0, 0, 594);
            Assert(Mm(-1261, -24.242131, 148.562476, 816.780091).ExtendsOutside(frame), "Pump PDF extents must warn");
            Assert(!Mm(-1241.0928, 5.6289, -394.5682, 242.0298).ExtendsOutside(frame), "Pump viewport 1 fits");
            Assert(!Mm(-1229.2327, 112.8681, -6.4444, 585.5144).ExtendsOutside(frame), "Pump viewport 2 fits");
            Assert(!Mm(-841, 0, 0, 594).ExtendsOutside(Mm(-841, 0, 0, 594)), "Algorithm frame fits");
            Assert(!Mm(-1261.5, -.5, .5, 594.5).ExtendsOutside(frame), "Ignore sub-millimeter bounds noise");
            Assert(Mm(-1263, 10, -1000, 20).ExtendsOutside(frame), "Left overflow");
            Assert(Mm(-100, -2, -10, 20).ExtendsOutside(frame), "Bottom overflow");
            Assert(Mm(-100, 10, 2, 20).ExtendsOutside(frame), "Right overflow");
            Assert(Mm(-100, 10, -10, 596).ExtendsOutside(frame), "Top overflow");
            Assert(!Mm(100, 200, 200, 300).ExtendsOutside(Mm(100, 200, 200, 300)), "Translated frame");

            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "results", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            string first = Path.Combine(folder, "01.pdf");
            string second = Path.Combine(folder, "02.pdf");
            string combined = Path.Combine(folder, "combined.pdf");
            string corrupt = Path.Combine(folder, "corrupt.pdf");
            MakePdf(first, "FIRST PRINTED SHEET");
            MakePdf(second, "SECOND PRINTED SHEET");
            Assert(PrintedPdfOutput.IsComplete(first), "Completed PDF detected");
            using (var locked = new FileStream(first, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                Assert(!PrintedPdfOutput.IsComplete(first), "Still-open printer file is not complete");
            File.WriteAllText(corrupt, "%PDF-1.4\nunfinished");
            Assert(!PrintedPdfOutput.IsComplete(corrupt), "Truncated PDF rejected");
            Assert(!PrintedPdfOutput.IsComplete(Path.Combine(folder, "missing.pdf")), "Missing PDF rejected");

            bool timedOut = false;
            try { PrintedPdfOutput.WaitForFiles(new[] { corrupt }, TimeSpan.FromMilliseconds(50)); }
            catch (TimeoutException) { timedOut = true; }
            Assert(timedOut, "Bounded wait reports pending file");
            string delayed = Path.Combine(folder, "delayed.pdf");
            var producer = Task.Run(() => { Thread.Sleep(350); File.Copy(second, delayed); });
            PrintedPdfOutput.WaitForFiles(new[] { first, delayed }, TimeSpan.FromSeconds(5));
            producer.Wait();
            Assert(PrintedPdfOutput.IsComplete(delayed), "Wait includes asynchronously completed printer job");
            PrintedPdfOutput.WaitForFiles(new string[0], TimeSpan.Zero);

            PrintedPdfOutput.Combine(new[] { first, second }, combined);
            using (var reader = new PdfReader(combined))
            {
                Assert(reader.NumberOfPages == 2, "Partial batch merges all completed pages");
                Assert(PdfTextExtractor.GetTextFromPage(reader, 1).Contains("FIRST"), "Order: first");
                Assert(PdfTextExtractor.GetTextFromPage(reader, 2).Contains("SECOND"), "Order: second");
            }
            Assert(File.Exists(first) && File.Exists(second), "Merge does not remove sources");
            byte[] previous = File.ReadAllBytes(combined);
            bool failed = false;
            try { PrintedPdfOutput.Combine(new[] { first, corrupt }, combined); }
            catch { failed = true; }
            Assert(failed, "Corrupt page aborts whole merge");
            Assert(previous.SequenceEqual(File.ReadAllBytes(combined)), "Failed merge preserves prior result");
            Assert(File.Exists(first) && File.Exists(corrupt), "Failed merge preserves all inputs");
            failed = false;
            try { PrintedPdfOutput.Combine(new string[0], combined); }
            catch (ArgumentException) { failed = true; }
            Assert(failed && previous.SequenceEqual(File.ReadAllBytes(combined)), "Zero printed pages preserves output");
            PrintedPdfOutput.Combine(new[] { second }, combined);
            using (var reader = new PdfReader(combined))
                Assert(reader.NumberOfPages == 1, "One-sheet partial batch is valid");
            Assert(Directory.GetFiles(folder, "*.pdf").Length == 5, "Temporary merge files cleaned up");
            Console.WriteLine("PASS: " + _checks + " regression checks. PDF fixtures: " + folder);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }
}
