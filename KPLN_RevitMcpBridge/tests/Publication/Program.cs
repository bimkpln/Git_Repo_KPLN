using System;
using System.Collections.Generic;
using KPLN_RevitMcpBridge.Services;
using KPLN_RevitMcpBridge.Server;

namespace Autodesk.Revit.DB { public class Document { public string Title => "Fixture"; } }
namespace KPLN_Publication
{
    public enum Processing { VectorProcessing, RasterProcessing }
    public enum Color { Color, Monochrome, GrayScale }
    public enum Quality { Low, Medium, High, Presentation }
    public class YayPrintSettings
    {
        public string printerName, outputPDFFolder, pdfNameConstructor;
        public Processing hiddenLineProcessing;
        public Color colorsType;
        public Quality rasterQuality;
        public bool isPDFExport, isPrintToPaper, isMergePdfs, isUseOrientation, isRefreshSchedules, isExcludeBorders, isDWGExport;
    }
    public static class BatchPrintService
    {
        public const int ContractVersion = 1;
        public static bool Fail;
        public static object PrintAllSheets(Autodesk.Revit.DB.Document doc, YayPrintSettings settings, bool dryRun, int timeout)
        {
            if (Fail) throw new InvalidOperationException("simulated printer error");
            return new Dictionary<string, object> { { "settings", settings }, { "dry_run", dryRun }, { "timeout", timeout } };
        }
    }
}
class Program
{
    static int checks;
    static void Check(bool value) { checks++; if (!value) throw new Exception("Check " + checks + " failed"); }
    static Dictionary<string, object> Input() => new Dictionary<string, object> { { "document_id", "doc" }, { "expected_revision", 7 } };
    static void Main()
    {
        var doc = new Autodesk.Revit.DB.Document();
        var result = (Dictionary<string, object>)PublicationPrintService.Print(doc, Input());
        var settings = (KPLN_Publication.YayPrintSettings)result["settings"];
        Check((bool)result["dry_run"]);
        Check((int)result["timeout"] == 30);
        Check((string)result["document_id"] == "doc");
        Check(Convert.ToInt64(result["revision_at_request"]) == 7);
        Check(settings.printerName == "PDFCreator");
        Check(settings.outputPDFFolder == @"C:\PDF_Print");
        Check(settings.pdfNameConstructor == "<Номер листа>_<Имя листа>.pdf");
        Check(settings.colorsType == KPLN_Publication.Color.Color);
        Check(settings.hiddenLineProcessing == KPLN_Publication.Processing.VectorProcessing);
        Check(settings.rasterQuality == KPLN_Publication.Quality.Medium);
        Check(settings.isRefreshSchedules && settings.isPDFExport);
        Check(!settings.isDWGExport && !settings.isPrintToPaper && !settings.isMergePdfs && !settings.isExcludeBorders && !settings.isUseOrientation);
        foreach (var invalid in new object[] { "wrong", new Dictionary<string, object> { { "unknown", true } },
            new Dictionary<string, object> { { "isRefreshSchedules", "true" } },
            new Dictionary<string, object> { { "colorsType", "missing" } } })
        {
            var input = Input(); input["settings"] = invalid;
            try { PublicationPrintService.Print(doc, input); throw new Exception("Invalid setting accepted"); }
            catch (BridgeException ex) { Check(ex.Message.Length > 0); }
        }
        KPLN_Publication.BatchPrintService.Fail = true;
        try { PublicationPrintService.Print(doc, Input()); throw new Exception("Failure swallowed"); }
        catch (BridgeException ex) { Check(ex.Message == "simulated printer error"); }
        Console.WriteLine("Publication adapter: " + checks + " checks passed.");
    }
}
