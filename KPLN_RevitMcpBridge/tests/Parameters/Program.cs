using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;
internal static class Program
{
    private static void Assert(bool yes, string message) { if (!yes) throw new Exception(message); }
    private static void Fails(Action action, string code)
    {
        try { action(); throw new Exception("Expected failure: " + code); }
        catch (BridgeException e) { Assert(e.Code == code, "Wrong failure: " + e.Code); }
    }
    private static Dictionary<string, object> Request(bool dryRun, params object[] updates) => Json.Parse(Json.Serialize(new { dry_run = dryRun, updates }));
    private static object Row(string id, object before, object after, string parameter = "1", string units = null) => new { element_unique_id = id, parameter_id = parameter, expected_value = before, value = after, value_units = units };
    private static void Main()
    {
        var originalCulture = Thread.CurrentThread.CurrentCulture;
        try
        {
            foreach (var culture in new[] { "en-US", "ru-RU" })
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                foreach (double number in new[] { 1.3123359580052643, -1.3123359580052643, 400.0 / 304.8, Math.PI, 0.1, 0.0, 1.0, 1e-30, 1e30 })
                    Assert(Json.Double(Json.Parse(Json.Serialize(new { value = number }))["value"]) == number,
                        "JSON double round trip failed: " + number.ToString("R", CultureInfo.InvariantCulture) + " / " + culture);
            }
        }
        finally { Thread.CurrentThread.CurrentCulture = originalCulture; }
        Assert(Json.Double(1) == 1.0 && Json.Double(1L) == 1.0, "integer JSON numbers");
        foreach (var invalid in new object[] { null, true, "1.3123359580052643", double.NaN, double.PositiveInfinity, double.NegativeInfinity })
            Fails(() => Json.Double(invalid), "invalid_value");
        var a = new Parameter { Id = new ElementId(1), StorageType = StorageType.String, Value = "old" };
        var b = new Parameter { Id = new ElementId(1), StorageType = StorageType.String, Value = "old" };
        var doc = new Document { Elements = new List<Element> { new Element { Id = new ElementId(11), UniqueId = "a", Parameters = new List<Parameter> { a } }, new Element { Id = new ElementId(12), UniqueId = "b", Parameters = new List<Parameter> { b } } } };
        ParameterService.Set(doc, Request(true, Row("a", "old", "new")));
        Assert((string)a.Value == "old", "preview wrote");
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "stale", "new"))), "stale_parameter");
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "old", "new"), Row("a", "old", "new"))), "duplicate_update");
        b.IsReadOnly = true;
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "old", "new"), Row("b", "old", "new"))), "read_only_parameter");
        Assert((string)a.Value == "old", "partial preflight mutation"); b.IsReadOnly = false;
        b.FailSet = true;
        try { ParameterService.Set(doc, Request(false, Row("a", "old", "new"), Row("b", "old", "new"))); throw new Exception("expected Set failure"); }
        catch (Exception e) { Assert(e.Message == "Set failed", "wrong set failure"); }
        Assert((string)a.Value == "old", "Set failure did not roll back previous write"); b.FailSet = false;
        doc.FailCommit = true;
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "old", "new"))), "transaction_rolled_back");
        Assert((string)a.Value == "old", "warning did not roll back"); doc.FailCommit = false;
        doc.OnCommit = () => a.Value = "updater changed";
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "old", "new"))), "verification_failed");
        Assert((string)a.Value == "old", "group did not undo commit updater"); doc.OnCommit = null;
        ParameterService.Set(doc, Request(false, Row("a", "old", "new"), Row("b", "old", "new")));
        Assert((string)a.Value == "new" && (string)b.Value == "new", "successful batch");
        a.StorageType = StorageType.Double; a.Value = 1.0;
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", 1.0, 2.0))), "invalid_input");
        ParameterService.Set(doc, Request(false, Row("a", 1.0, 2.0, units: "revit_internal")));
        Assert((double)a.Value == 2.0, "double units");
        const double precise = 1.3123359580052643;
        a.Value = precise;
        var preciseRequest = Request(true, Row("a", precise, precise, units: "revit_internal"));
        ParameterService.Set(doc, preciseRequest);
        Assert((double)a.Value == precise, "double preview wrote");
        preciseRequest["dry_run"] = false;
        ParameterService.Set(doc, preciseRequest);
        Assert((double)a.Value == precise, "double round trip changed the value");
        double adjacent = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(precise) + 1);
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", adjacent, 2.0, units: "revit_internal"))), "stale_parameter");
        Assert((double)a.Value == precise, "stale double changed the value");
        a.Value = 2.0;
        ParameterService.Set(doc, Request(false, Row("a", 2.0, precise, units: "revit_internal")));
        Assert((double)a.Value == precise, "new double was rounded during parsing");
        a.StorageType = StorageType.Integer; a.Value = 1;
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", 1, 1.5))), "invalid_input");
        a.StorageType = StorageType.String; a.Value = null;
        ParameterService.Set(doc, Request(false, Row("a", null, "filled")));
        Assert((string)a.Value == "filled", "unset value handling");
        doc.IsFamilyDocument = true;
        Fails(() => ParameterService.Set(doc, Request(false, Row("a", "filled", "new"))), "family_write_not_supported");
        Console.WriteLine("PASS: preview, stale/duplicate/read-only checks, atomic Set failure, warning rollback, post-commit updater rollback, success, units, null, family restriction (API doubles).");
    }
}
