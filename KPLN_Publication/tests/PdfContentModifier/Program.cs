using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using iTextSharp.text.pdf;
using KPLN_Publication.PdfWorker;

class Program
{
    static void Main(string[] args)
    {
        ParserRegression.Run();
        BorderHideRegression.Run();
        if (args.Length != 2) return;
        Directory.CreateDirectory(args[1]);
        for (int mode = 0; mode < 4; mode++)
        {
            using (var reader = new PdfReader(args[0]))
            using (var output = File.Create(Path.Combine(args[1], "mode" + mode + ".pdf")))
            using (var stamper = new PdfStamper(reader, output))
            {
                var borderConverter = new PdfContentToBlackWhiteConverter();
                var colorConverter = new PdfContentToBlackWhiteConverter();
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    if ((mode & 1) != 0)
                        borderConverter.ConvertWithHide(reader, page, new List<Color> { Color.FromArgb(3, 2, 51) });
                    if ((mode & 2) != 0)
                        colorConverter.Convert(reader, page, new List<Color> { Color.Blue });
                }
            }
            Console.WriteLine("PASS real PDF, mode=" + mode);
        }
    }
}
