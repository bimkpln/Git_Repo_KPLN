using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;
class Program
{
    static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    static Dictionary<string, object> Map(object value) => Json.Parse(Json.Serialize(value));
    static void Main()
    {
        var doc = new Document();
        doc.Elements[2] = new View();
        doc.Elements[3] = new FamilySymbol { FamilyName = "Выноска", Parameters = new[] { "Параметр типа" } };
        var text = new TextNote { Text = "П2\r\n600×400 — 12 м³/ч", Coord = new XYZ { X = 1 }, ThrowBounds = true };
        var result = Map(AnnotationContentService.Read(doc, text));
        Check((string)result["text"] == text.Text, "multiline Unicode text lost");
        Check((string)result["kind"] == "text_note", "wrong kind");
        Check((string)Map(result["bounds"])["status"] == "unavailable", "optional bounds failure lost text");
        Check(AnnotationContentService.Read(doc, new Element()) == null, "non-annotation included");
        var generic = new FamilyInstance { Category = new Category { Id = new ElementId(-2000150) }, Parameters = new[] { "Марка вручную", "142" } };
        result = Map(AnnotationContentService.Read(doc, generic));
        Check((string)result["kind"] == "generic_annotation", "generic missing");
        Check((string)result["family_name"] == "Выноска", "family name missing");
        Check(result["parameters"] != null && result["type_parameters"] != null, "parameter scope lost");
        Check(result["text"] == null && (string)result["text_status"] == "not_resolved_from_family_labels", "parameters falsely treated as rendered text");
        doc.Elements.Remove(3);
        result = Map(AnnotationContentService.Read(doc, generic));
        Check(result["type"] == null && result["type_parameters"] == null, "missing type failed");
        generic.Location = new LocationPoint { Unsupported = true, Point = new XYZ { X = 1 } };
        generic.Box = new BoundingBoxXYZ();
        result = Map(AnnotationContentService.Read(doc, generic));
        var location = Map(result["location"]);
        Check(location["rotation_radians"] == null && (string)location["rotation_status"] == "not_supported", "unsupported rotation broke annotation read");
        Check((string)location["rotation_error"] == "Unable to extract rotation from family instance.", "rotation reason lost");
        Check((string)Map(result["bounds"])["status"] == "available", "view bounds missing");
        Console.WriteLine("PASS: TextNote multiline/Unicode, generic instance/type values, missing type, optional bounds failure, explicit label limitations.");
    }
}
