using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using KPLN_Publication.PdfWorker;

static class BorderHideRegression
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static string Text(byte[] bytes) { return Encoding.ASCII.GetString(bytes); }
    static byte[] Bytes(string value) { return Encoding.ASCII.GetBytes(value); }

    public static void Run()
    {
        byte[] pdf;
        using (var memory = new MemoryStream())
        {
            var document = new Document();
            var writer = PdfWriter.GetInstance(document, memory);
            document.Open();
            writer.DirectContent.SetLiteral("0 1 0 rg 1 0 0 RG 0.1 0.2 0.3 0.4 k 0.4 0.3 0.2 0.1 K /DeviceRGB cs 0 1 0 sc /DeviceRGB CS 1 0 0 SC 0 0 1 scn 1 0 1 SCN 0.2 g 0.3 G\n");
            writer.DirectContent.SetLiteral("0.011765 0.007843 0.2 RG 0 0 m 10 10 l S\n");
            var template = writer.DirectContent.CreateTemplate(20, 20);
            template.SetLiteral("0 1 0 rg 0.011765 0.007843 0.2 RG 0 0 10 10 re B\n");
            writer.DirectContent.AddTemplate(template, 0, 0);
            var image = iTextSharp.text.Image.GetInstance(1, 1, 3, 8, new byte[] { 255, 0, 0 });
            image.SetAbsolutePosition(0, 0);
            writer.DirectContent.AddImage(image);
            document.Close();
            pdf = memory.ToArray();
        }

        using (var reader = new PdfReader(pdf))
        {
            string before = Text(reader.GetPageContent(1));
            var resources = reader.GetPageN(1).GetAsDict(PdfName.RESOURCES);
            var xobjects = resources.GetAsDict(PdfName.XOBJECT);
            var images = new Dictionary<PdfName, byte[]>();
            foreach (PdfName key in xobjects.Keys)
            {
                var stream = xobjects.GetDirectObject(key) as PRStream;
                if (stream.GetAsName(PdfName.SUBTYPE) == PdfName.IMAGE)
                    images.Add(key, PdfReader.GetStreamBytes(stream));
            }

            var converter = new PdfContentToBlackWhiteConverter();
            converter.ConvertWithHide(reader, 1, new List<Color> { Color.FromArgb(3, 2, 51) });
            string after = Text(reader.GetPageContent(1));
            string expected = Text(new PdfContentModifier().Modify(Bytes(before.Replace("0.011765 0.007843 0.2 RG", "1 1 1 RG")), resources));
            Check(after == expected, "Border hiding changed unrelated page colors or operators");
            foreach (PdfName key in xobjects.Keys)
            {
                var stream = xobjects.GetDirectObject(key) as PRStream;
                if (stream.GetAsName(PdfName.SUBTYPE) == PdfName.FORM)
                    Check(Text(PdfReader.GetStreamBytes(stream)).Contains("0 1 0 rg 1 1 1 RG"), "Nested form colors/border incorrect");
                if (stream.GetAsName(PdfName.SUBTYPE) == PdfName.IMAGE)
                    Check(images[key].SequenceEqual(PdfReader.GetStreamBytes(stream)), "Image recolored");
            }
            Check(images.Count == 1, "Missing image fixture");
            // Старый обработчик изображений монохромного режима пишет диагностические BMP вне теста.
            // Сохранность IMAGE уже проверена; далее проверяем только совместную обработку векторов.
            foreach (PdfName key in images.Keys)
                xobjects.Remove(key);
            new PdfContentToBlackWhiteConverter().Convert(reader, 1, new List<Color> { Color.Red });
            Check(Text(reader.GetPageContent(1)).Contains("0 0 0 rg"), "Monochrome mode stopped converting fills");
            Check(Text(reader.GetPageContent(1)).Contains("1 1 1 RG"), "Combined mode restored the hidden border");
            Check(Text(reader.GetPageContent(1)).Contains("1 0 0 RG"), "Excluded red stroke was changed");
            foreach (PdfName key in xobjects.Keys)
            {
                var stream = xobjects.GetDirectObject(key) as PRStream;
                if (stream.GetAsName(PdfName.SUBTYPE) == PdfName.FORM)
                {
                    Check(Text(PdfReader.GetStreamBytes(stream)).Contains("0 0 0 rg"), "Mode switch skipped nested form");
                    Check(Text(PdfReader.GetStreamBytes(stream)).Contains("1 1 1 RG"), "Combined mode restored nested border");
                }
            }
        }
        Console.WriteLine("PASS: hide only border; preserve RGB/CMYK/general colors, fills and image; process nested forms and mode switch");

        using (var reader = new PdfReader(pdf))
        {
            byte[] before = reader.GetPageContent(1);
            new PdfContentToBlackWhiteConverter().ConvertWithHide(reader, 1, new List<Color>());
            var normalized = new PdfContentModifier().Modify(before, reader.GetPageN(1).GetAsDict(PdfName.RESOURCES));
            Check(normalized.SequenceEqual(reader.GetPageContent(1)), "Empty hide list changed colors");
        }
        Console.WriteLine("PASS: empty hide list preserves colors");

        using (var reader = new PdfReader(pdf))
        {
            var modifier = new PdfContentModifier();
            int calls = 0;
            modifier.RegisterOperator("rg", (op, args) => { calls++; return args; });
            modifier.Modify(reader, 1, (op, args) => false);
            Check(calls == 0, "Filtered handler was called");
            modifier.Modify(reader, 1);
            Check(calls > 0, "Filter leaked into next page pass");
            int previous = calls;
            modifier.Modify(reader, 1, (op, args) => { throw new Exception("Unsupported resource"); });
            Check(calls == previous, "Throwing filter invoked handler");
            modifier.Modify(reader, 1);
            Check(calls > previous, "Throwing filter leaked state");
        }
        Console.WriteLine("PASS: filter scope and unsupported-resource fallback");

        using (var reader = new PdfReader(pdf))
        {
            reader.SetPageContent(1, Bytes("0 0 1 rg 0 1 0 rg 1 1 1 rg 0.5 0.5 0.5 rg 0 0 1 RG 0 1 0 RG 1 1 1 RG"));
            new PdfContentToBlackWhiteConverter().Convert(reader, 1, new List<Color> { Color.Blue });
            Check(Text(reader.GetPageContent(1)).Trim() == "0 0 1 rg 0 0 0 rg 1 1 1 rg 0.5 0.5 0.5 rg 0 0 1 RG 0 0 0 RG 1 1 1 RG", "Blue exception or neutral colors changed");
        }
        Console.WriteLine("PASS: exact blue exception, white and grey preserved; other RGB changed");
    }
}
