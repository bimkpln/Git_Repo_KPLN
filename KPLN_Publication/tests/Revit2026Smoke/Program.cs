using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using KPLN_Publication;
using KPLN_Publication.PdfWorker;

internal static class Program
{
    private static void Main()
    {
        string folder = Path.Combine(AppContext.BaseDirectory, "fixtures", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string input = Path.Combine(folder, "input.pdf");
        using (var stream = File.Create(input))
        using (var document = new Document())
        {
            var writer = PdfWriter.GetInstance(document, stream);
            document.Open();
            document.Add(new Paragraph("Publication .NET 8 smoke test"));
            writer.DirectContent.SetColorFill(BaseColor.RED);
            writer.DirectContent.Rectangle(50, 50, 100, 100);
            writer.DirectContent.Fill();
            using (var bitmap = new Bitmap(4, 4))
            {
                using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.Blue);
                var image = iTextSharp.text.Image.GetInstance(bitmap, ImageFormat.Png);
                document.Add(image);
            }
        }

        string gray = Path.Combine(folder, "gray.pdf");
        PdfWorker.ConvertToGrayScale(input, gray);
        string borders = Path.Combine(folder, "borders.pdf");
        PdfWorker.SetHideColors(new List<PdfColor>());
        PdfWorker.ConvertToBorderHide(input, borders);
        foreach (string path in new[] { gray, borders })
            using (var reader = new PdfReader(path))
                if (reader.NumberOfPages != 1 || reader.GetPageContent(1).Length == 0)
                    throw new Exception("Invalid converted PDF: " + path);

        var messages = new List<string>();
        var logger = new Logger(messages.Add);
        string merged = Path.Combine(folder, "merged.pdf");
        PdfWorker.CombineMultiplyPDFs(new List<string> { gray, borders }, merged, logger);
        using (var reader = new PdfReader(merged))
            if (reader.NumberOfPages != 2) throw new Exception("PDF merge failed");
        if (messages.Count == 0) throw new Exception("Callback logging failed");

        new Logger().Write("Revit2026 file logger smoke test");
        string logs = Path.Combine(AppContext.BaseDirectory, "logs");
        bool logged = false;
        foreach (string path in Directory.GetFiles(logs, "*.log"))
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
                logged |= reader.ReadToEnd().Contains("Revit2026 file logger smoke test");
        }
        if (!logged) throw new Exception("File logging failed");
        Console.WriteLine("PASS: .NET 8 PDF creation, vector/raster conversion, border processing, merge, callback/file logging.");
    }
}
