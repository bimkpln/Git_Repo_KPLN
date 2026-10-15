using System;
using System.Collections.Generic;
using System.Linq;
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
    private static object Value(long id, object value, string units = null) => new { parameter_id = id.ToString(), value, value_units = units };
    private static object ValueSet(long id, object expectedValue, object value, string units = null) => new { parameter_id = id.ToString(), expected_value = expectedValue, value, value_units = units };
    private static Dictionary<string, object> Request(bool dryRun, params object[] types) => Json.Parse(Json.Serialize(new { dry_run = dryRun, base_type_name = "base", types }));
    private static Dictionary<string, object> UpdateRequest(bool dryRun, params object[] types) => Json.Parse(Json.Serialize(new { dry_run = dryRun, types }));
    private static object Type(string name, params object[] values) => new { name, values };

    private static void Main()
    {
        CheckAssociations();

        var text = new FamilyParameter { Id = new ElementId(1), Definition = new Definition("Mark"), StorageType = StorageType.String };
        var length = new FamilyParameter { Id = new ElementId(2), Definition = new Definition("Length"), StorageType = StorageType.Double };
        var nested = new FamilyParameter { Id = new ElementId(3), Definition = new Definition("Nested"), StorageType = StorageType.ElementId };
        var formula = new FamilyParameter { Id = new ElementId(4), Definition = new Definition("Formula"), StorageType = StorageType.Double, Formula = "Length" };
        var baseType = new FamilyType { Name = "base" };
        baseType.Values[text] = "base"; baseType.Values[length] = 1.0; baseType.Values[nested] = new ElementId(900); baseType.Values[formula] = 1.0;
        var doc = new Document { Elements = new List<Element> { new Element { Id = new ElementId(900) } } };
        doc.FamilyManager.ParameterList.AddRange(new[] { text, length, nested, formula });
        doc.FamilyManager.TypeList.Add(baseType); doc.FamilyManager.CurrentType = baseType;

        var request = Request(true,
            Type("A", Value(1, "A"), Value(2, 2.0, "revit_internal"), Value(3, "900")),
            Type("B", Value(1, "B"), Value(2, 3.0, "revit_internal")));
        FamilyTypeService.Create(doc, request);
        Assert(doc.FamilyManager.TypeList.Count == 1, "preview created a type");
        request["dry_run"] = false;
        FamilyTypeService.Create(doc, request);
        Assert(doc.FamilyManager.TypeList.Count == 3, "types were not created");
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "A").AsDouble(length) == 2.0, "A value");
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "B").AsString(text) == "B", "B value");
        Assert(baseType.AsString(text) == "base", "base type changed");

        Fails(() => FamilyTypeService.Create(doc, Request(true, Type("A", Value(1, "x")))), "type_exists");
        Fails(() => FamilyTypeService.Create(doc, Request(true, Type("C", Value(2, 4.0)))), "invalid_units");
        Fails(() => FamilyTypeService.Create(doc, Request(true, Type("C", Value(4, 4.0, "revit_internal")))), "read_only_parameter");

        var currentBeforeUpdate = doc.FamilyManager.CurrentType;
        var update = UpdateRequest(true, Type("A", ValueSet(1, "A", "A updated"), ValueSet(2, 2.0, 2.5, "revit_internal")));
        FamilyTypeService.Update(doc, update);
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "A").AsString(text) == "A", "preview updated a value");
        update["dry_run"] = false;
        FamilyTypeService.Update(doc, update);
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "A").AsString(text) == "A updated", "text value was not updated");
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "A").AsDouble(length) == 2.5, "double value was not updated");
        Assert(object.ReferenceEquals(doc.FamilyManager.CurrentType, currentBeforeUpdate), "current type was not restored");

        const double precise = 1.3123359580052643;
        var typeA = doc.FamilyManager.TypeList.Single(t => t.Name == "A");
        typeA.Values[length] = precise;
        var preciseUpdate = UpdateRequest(true, Type("A", ValueSet(2, precise, precise, "revit_internal")));
        FamilyTypeService.Update(doc, preciseUpdate);
        Assert(typeA.AsDouble(length) == precise, "double preview wrote");
        preciseUpdate["dry_run"] = false;
        FamilyTypeService.Update(doc, preciseUpdate);
        Assert(typeA.AsDouble(length) == precise, "double round trip changed the value");
        double adjacent = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(precise) + 1);
        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(false, Type("A", ValueSet(2, adjacent, 2.5, "revit_internal")))), "stale_value");
        Assert(typeA.AsDouble(length) == precise, "stale double changed the value");
        typeA.Values[length] = 2.5;
        FamilyTypeService.Update(doc, UpdateRequest(false, Type("A", ValueSet(2, 2.5, precise, "revit_internal"))));
        Assert(typeA.AsDouble(length) == precise, "new family double was rounded during parsing");
        FamilyTypeService.Create(doc, Request(false, Type("Precise", Value(2, precise, "revit_internal"))));
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "Precise").AsDouble(length) == precise, "created family double was rounded during parsing");
        typeA.Values[length] = 2.5;

        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(true, Type("A", ValueSet(1, "A", "stale")))), "stale_value");
        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(true, Type("missing", ValueSet(1, "x", "y")))), "not_found");
        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(true, Type("A", ValueSet(2, 2.5, 3.0)))), "invalid_units");
        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(true, Type("A", ValueSet(4, 1.0, 2.0, "revit_internal")))), "read_only_parameter");

        doc.FailCommit = true;
        Fails(() => FamilyTypeService.Create(doc, Request(false, Type("C", Value(1, "C")))), "transaction_rolled_back");
        Assert(!doc.FamilyManager.TypeList.Any(t => t.Name == "C"), "warning did not roll back created type");
        Fails(() => FamilyTypeService.Update(doc, UpdateRequest(false, Type("A", ValueSet(1, "A updated", "rolled back")))), "transaction_rolled_back");
        Assert(doc.FamilyManager.TypeList.Single(t => t.Name == "A").AsString(text) == "A updated", "warning did not roll back updated value");
        Console.WriteLine("PASS: family type preview, atomic creation and update, stale checks, values, validation and warning rollback (API doubles).");
    }

    private static void CheckAssociations()
    {
        var glass = new FamilyParameter { Id = new ElementId(2865054), Definition = new Definition("Остекление_Материал"), StorageType = StorageType.ElementId };
        var typo = new FamilyParameter { Id = new ElementId(2958061), Definition = new Definition("Остекление_Матераил"), StorageType = StorageType.ElementId };
        var symbol = new FamilySymbol { Id = new ElementId(3000000001), UniqueId = "leaf-type", Name = "Glazed", FamilyName = "Nested leaf" };
        var first = new FamilyInstance { Id = new ElementId(301), UniqueId = "first-leaf", Name = "Glazed", Symbol = symbol };
        var second = new FamilyInstance { Id = new ElementId(302), UniqueId = "second-leaf", Name = "Glazed", Symbol = symbol };
        glass.AssociatedParameters.Add(new Parameter { Id = new ElementId(50), Definition = new Definition("Glass"), StorageType = StorageType.ElementId, Element = second });
        glass.AssociatedParameters.Add(new Parameter { Id = new ElementId(50), Definition = new Definition("Glass"), StorageType = StorageType.ElementId, Element = first });
        typo.AssociatedParameters.Add(new Parameter { Id = new ElementId(51), Definition = new Definition("Material"), StorageType = StorageType.ElementId, Element = symbol });

        var result = Json.Parse(Json.Serialize(FamilyTypeService.DescribeParameter(glass)));
        var rows = result.List("associated_parameters");
        Assert(rows.Count == 2 && Json.Integer(result.Get("associated_parameters_count"), "count") == 2, "association count");
        var row = (Dictionary<string, object>)rows[0];
        Assert(((Dictionary<string, object>)row["element"]).Text("unique_id") == "first-leaf", "association owner/order");
        Assert(row.Text("parameter_name") == "Glass" && row.Text("family_name") == "Nested leaf", "association names");
        Assert(!result.Flag("associated_parameters_truncated"), "unexpected truncation");

        result = Json.Parse(Json.Serialize(FamilyTypeService.DescribeParameter(typo)));
        row = (Dictionary<string, object>)result.List("associated_parameters")[0];
        var element = (Dictionary<string, object>)row["element"];
        Assert(element.Text("id") == "3000000001" && element.Flag("is_type"), "type owner / 64-bit id");
        Assert(result.Text("name") == "Остекление_Матераил", "similarly named parameters merged");

        typo.AssociatedParameters.Clear();
        result = Json.Parse(Json.Serialize(FamilyTypeService.DescribeParameter(typo)));
        Assert(((System.Collections.IList)result["associated_parameters"]).Count == 0, "unbound parameter");
        for (int i = 0; i < 201; i++) typo.AssociatedParameters.Add(glass.AssociatedParameters[0]);
        result = Json.Parse(Json.Serialize(FamilyTypeService.DescribeParameter(typo)));
        Assert(result.List("associated_parameters").Count == 200 && result.Flag("associated_parameters_truncated"), "bounded associations");
        Assert(Json.Integer(result.Get("associated_parameters_count"), "count") == 201, "full association count");
        Console.WriteLine("PASS: family parameter associations, nested instances/types, distinct spellings, 64-bit IDs, empty and truncated results.");
    }
}
