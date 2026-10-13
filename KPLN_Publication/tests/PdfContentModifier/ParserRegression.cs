using System;
using System.Collections.Generic;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using KPLN_Publication.PdfWorker;

class ParserRegression
{
    static byte[] Bytes(string value) { return Encoding.ASCII.GetBytes(value); }
    static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    public static void Run()
    {

        var modifier = new PdfContentModifier();
        modifier.RegisterOperator("rg", (op, args) => new List<PdfObject> {
            new PdfNumber(0), new PdfNumber(0), new PdfNumber(0), op });
        var input = Bytes("q 1 0 0 rg BI /W 1 /H 1 /BPC 8 /CS /RGB ID >(< EI 0 1 0 rg Q");
        var result = modifier.Modify(input, new PdfDictionary());
        string text = Encoding.ASCII.GetString(result);
        Check(text.Contains("ID >(< EI"), "Inline image bytes were changed");
        Check(text.Contains("0 0 0 rg Q"), "Operators after image were not processed");
        Check(text.Contains("q 0 0 0 rg"), "Operators before image were not processed");
        var parser = new PdfContentParser(new PRTokeniser(new RandomAccessFileOrArray(result)));
        var operands = new List<PdfObject>();
        int images = 0;
        while (parser.Parse(operands).Count > 0)
            if (operands[operands.Count - 1].ToString() == "BI") {
                InlineImageUtils.ParseInlineImage(parser, null); images++;
            }
        Check(images == 1, "Output image cannot be parsed");
        Console.WriteLine("PASS: binary inline image preserved; surrounding colors converted; output parses");

        var outerResources = new PdfDictionary();
        outerResources.Put(new PdfName("Marker"), new PdfNumber(7));
        modifier.RegisterOperator("Do", (op, args) => {
            modifier.Modify(Bytes(">"), new PdfDictionary());
            return args;
        });
        modifier.RegisterOperator("Q", (op, args) => {
            Check(modifier.ResourceDictionary.GetAsNumber(new PdfName("Marker")).IntValue == 7,
                "Nested failure leaked resources");
            return args;
        });
        result = modifier.Modify(Bytes("q /Broken Do 1 0 0 rg Q"), outerResources);
        Check(Encoding.ASCII.GetString(result).Contains("q /Broken Do 0 0 0 rg Q"), "Nested failure lost page operators");
        Check(modifier.ResourceDictionary.GetDirectObject(new PdfName("Marker")) == null, "Resources leaked after conversion");
        Console.WriteLine("PASS: nested parse failure does not corrupt parent content or resource stack");

        bool failed = false;
        try { modifier.Modify(Bytes(">"), outerResources); } catch { failed = true; }
        Check(failed, "Malformed PDF was silently accepted");
        Check(modifier.ResourceDictionary.GetDirectObject(new PdfName("Marker")) == null, "Resources leaked after failure");
        Console.WriteLine("PASS: malformed content raises an error and cleans up state");
    }
}
