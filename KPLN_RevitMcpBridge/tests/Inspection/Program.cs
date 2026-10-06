using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

internal static class Program
{
    private static int _checks;
    private static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
        _checks++;
    }

    private static Dictionary<string, object> Map(object value) => Json.Parse(Json.Serialize(value));
    private static IList Items(object value) => (IList)value;
    private static Dictionary<string, object> Input(params object[] pairs)
    {
        var result = new Dictionary<string, object>();
        for (int i = 0; i < pairs.Length; i += 2) result.Add((string)pairs[i], pairs[i + 1]);
        return result;
    }

    private static void Reject(Action action, string code)
    {
        try { action(); }
        catch (BridgeException ex) { Assert(ex.Code == code, "unexpected error: " + ex.Code); return; }
        throw new Exception("missing error: " + code);
    }

    private static void Main()
    {
        var unsupportedPoint = new LocationPoint { Unsupported = true, Point = new XYZ { X = 1, Y = 2, Z = 3 } };
        var point = Map(ElementLocationService.Read(new Group { Location = unsupportedPoint }));
        Assert(point["rotation_radians"] == null && (string)point["rotation_status"] == "not_supported", "unsupported rotation must not fail or become zero");
        Assert(Convert.ToDouble(Items(point["point_mm"])[0]) == 304.8, "point units");
        var ordinary = Map(ElementLocationService.Read(new Element { Location = new LocationPoint { Angle = 0.5 } }));
        Assert(Convert.ToDouble(ordinary["rotation_radians"]) == 0.5, "supported rotation retained");
        Assert(ElementLocationService.Read(new Element()) == null, "missing location");
        var curve = Map(ElementLocationService.Read(new Element { Location = new LocationCurve { Curve = new Curve() } }));
        Assert(Convert.ToDouble(curve["length_mm"]) == 304.8, "curve units retained");


        var category = new Category { Id = new ElementId(-10) };
        var doc = new Document();
        doc.Categories.Add(category);
        var a = new Element { Id = new ElementId(1), UniqueId = "a", Category = category };
        var b = new Element { Id = new ElementId(2), UniqueId = "b", Category = category };
        var c = new Group { Id = new ElementId(3), UniqueId = "nested", Category = category };
        var group = new Group
        {
            Id = new ElementId(4), UniqueId = "group", Category = new Category { Id = new ElementId(-1) },
            Members = new List<ElementId> { b.Id, c.Id, a.Id }, AttachedTypes = new List<ElementId> { new ElementId(20) }
        };
        doc.Elements.AddRange(new Element[] { a, b, c, group });
        var members = Map(ModelInspectionService.GroupMembers(doc, Input("group_unique_id", "group", "offset", 1, "limit", 1)));
        var member = Map(Items(members["items"])[0]);
        Assert((int)members["total_count"] == 3 && (int)members["next_offset"] == 2, "group pagination");
        Assert((int)member["member_index"] == 1 && (bool)Map(member["data"])["is_nested_group"], "group preserves direct member order");
        Assert(Items(members["available_attached_detail_group_type_ids"]).Count == 1, "attached type ids");
        Assert(Items(Map(ModelInspectionService.GroupMembers(doc, Input("group_unique_id", "group", "offset", 3)))["items"]).Count == 0, "group end page");
        Reject(() => ModelInspectionService.GroupMembers(doc, Input("group_unique_id", "a")), "wrong_element_type");


        var view = new View { Id = new ElementId(10), UniqueId = "view", Candidates = new HashSet<ElementId> { a.Id } };
        var sheet = new ViewSheet { Id = new ElementId(11), UniqueId = "sheet" };
        var viewport = new Viewport { Id = new ElementId(12), UniqueId = "vp", ViewId = view.Id };
        sheet.Viewports.Add(viewport.Id);
        var schedule = new ViewSchedule { Id = new ElementId(13), UniqueId = "schedule" };
        var placement = new ScheduleSheetInstance { Id = new ElementId(14), UniqueId = "placement", ScheduleId = schedule.Id, OwnerViewId = sheet.Id };
        doc.Elements.AddRange(new Element[] { view, sheet, viewport, schedule, placement });
        var contents = Map(ModelInspectionService.SheetContents(doc, Input("sheet_unique_id", "sheet")));
        Assert(Items(contents["viewports"]).Count == 1 && Items(contents["schedules"]).Count == 1, "sheet placements");
        Assert((int)Map(Map(Items(contents["viewports"])[0])["view"])["scale"] == 100, "placed view scale");
        var visible = Map(ModelInspectionService.ViewElements(doc, Input("view_unique_id", "view")));
        Assert(Items(visible["items"]).Count == 1 && ((string)visible["visibility_semantics"]).Contains("not_rendered"), "view collector must not claim rendered visibility");
        b.Hidden = true;
        var visibility = Map(ModelInspectionService.ViewVisibility(doc, Input("view_unique_id", "view", "unique_ids", new[] { "a", "b" })));
        Assert((bool)Map(Items(visibility["elements"])[0])["potentially_visible"], "candidate visibility");
        Assert(!(bool)Map(Items(visibility["elements"])[1])["potentially_visible"] && (bool)Map(Items(visibility["elements"])[1])["individually_hidden"], "hidden visibility evidence");
        view.IsTemplate = true;
        Reject(() => ModelInspectionService.ViewElements(doc, Input("view_unique_id", "view")), "unsupported_view");
        view.IsTemplate = false;


        schedule.Table.Body = new TableSectionData { FirstRowNumber = 4, FirstColumnNumber = 2, NumberOfRows = 3, NumberOfColumns = 3 };
        schedule.Definition.Fields.Add(new ScheduleField());
        schedule.Definition.Sorting.Add(new ScheduleSortGroupField { ShowFooter = true });
        schedule.Definition.Filters.Add(new ScheduleFilter { Value = 0.5 });
        var table = Map(ModelInspectionService.ScheduleData(doc, Input("schedule_unique_id", "schedule", "offset", 1, "limit", 1, "column_offset", 1, "column_limit", 1)));
        var row = Map(Items(table["rows"])[0]);
        Assert((int)row["row_index"] == 5 && (string)Map(Items(row["cells"])[0])["text"] == "Body:5:3", "nonzero Revit table indices");
        Assert((int)table["next_offset"] == 2 && (int)table["next_column_offset"] == 2, "two-dimensional pagination");
        Assert((string)Map(Map(Items(table["filters"])[0])["value"])["kind"] == "double_revit_internal", "filter units");
        Assert((string)Map(Items(table["fields"])[0])["display_type"] == "Totals", "field totals");
        Assert((bool)Map(Items(table["sorting_grouping"])[0])["show_footer"], "group footer");
        var emptyTable = Map(ModelInspectionService.ScheduleData(doc, Input("schedule_unique_id", "schedule", "section", "header")));
        Assert(Items(emptyTable["rows"]).Count == 0 && emptyTable["next_offset"] == null, "empty section no cell calls");
        Reject(() => ModelInspectionService.ScheduleData(doc, Input("schedule_unique_id", "schedule", "column_limit", 51)), "invalid_input");
        Reject(() => ModelInspectionService.ScheduleData(doc, Input("schedule_unique_id", "schedule", "section", "wrong")), "invalid_input");


        // Весь документ: иная категория, скрытый член группы, вложенная геометрия.
        doc = new Document();
        a.Geometry = new GeometryElement { Items = new List<GeometryObject> { new GeometryInstance { Nested = new GeometryElement { Items = new List<GeometryObject> { new Solid() } } } } };
        b.Category = new Category { Id = new ElementId(-12345) };
        b.Hidden = true;
        b.GroupId = new ElementId(3);
        var third = new Element { Id = new ElementId(5), UniqueId = "third", Category = category };
        doc.Elements.AddRange(new Element[] { a, b, c, group, third });
        var annotation = new Element { Id = new ElementId(40), UniqueId = "annotation", ViewSpecific = true, Category = category };
        var symbol = new ElementType { Id = new ElementId(41), UniqueId = "symbol", Category = category };
        doc.Elements.AddRange(new Element[] { annotation, symbol });
        a.Intersections.Add(annotation.Id);
        a.Intersections.Add(symbol.Id);
        var request = Input();
        foreach (var excludedCategory in new[] { BuiltInCategory.OST_Rooms, BuiltInCategory.OST_Cameras })
        {
            doc.Elements.Add(new Element
            {
                Id = new ElementId(50 + doc.Elements.Count), UniqueId = excludedCategory.ToString(),
                Category = new Category { Id = new ElementId((int)excludedCategory) },
                ThrowGeometry = true, Supported = false
            });
        }
        var result = Map(InterferenceService.Check(doc, request));
        Assert((int)Map(result["excluded_counts"])["rooms_excluded_by_policy"] == 1, "rooms excluded before geometry access");
        Assert((int)Map(result["excluded_counts"])["cameras_excluded_by_policy"] == 1, "cameras excluded before geometry access");
        Assert((string)result["status"] == "no_intersections" && (bool)result["complete"], "complete no-intersection result");
        Assert(!(bool)result["native_dialog_executed"] && (int)result["pair_count"] == 0, "API provenance and self-pair exclusion");
        a.Intersections.Add(b.Id);
        b.Intersections.Add(a.Id);
        result = Map(InterferenceService.Check(doc, request));
        Assert((string)result["status"] == "intersections_found" && (int)result["pair_count"] == 1, "symmetric pair deduplication");
        Assert((int)result["solid_element_count"] == 3, "all categories, hidden/group members; containers excluded");
        a.Intersections.Add(third.Id);
        third.Intersections.Add(a.Id);
        request["max_pairs"] = 1;
        result = Map(InterferenceService.Check(doc, request));
        Assert(!(bool)result["complete"] && (string)result["stop_reason"] == "pair_limit" && result["message"] == null, "truncation never claims clear");
        request.Remove("max_pairs");
        a.Intersections.Clear(); b.Intersections.Clear(); third.Intersections.Clear();
        b.Supported = false;
        result = Map(InterferenceService.Check(doc, request));
        Assert((string)result["status"] == "incomplete" && Items(result["unsupported_elements"]).Count == 1, "unsupported does not produce false clear");
        b.Supported = true;
        a.ThrowIntersection = true;
        result = Map(InterferenceService.Check(doc, request));
        Assert((string)result["status"] == "incomplete" && Items(result["errors"]).Count == 1, "API failure does not produce false clear");
        a.ThrowIntersection = false;
        Assert((string)Map(InterferenceService.Check(new Document(), request))["status"] == "empty_scope", "empty scope is not pass");
        request["right_category_ids"] = new[] { "-999" };
        Reject(() => InterferenceService.Check(doc, request), "invalid_input");
        request.Remove("right_category_ids");
        b.Geometry.Items.Add(new Mesh());
        result = Map(InterferenceService.Check(doc, request));
        Assert((bool)result["complete"], "test model policy ignores non-volume part");
        var details = Map(Map(Items(result["geometry_notes"])[0])["geometry_details"]);
        Assert((int)details["mesh_count"] == 1 && (int)details["volume_solid_count"] == 1, "mixed geometry diagnostics");
        Assert((int)result["solid_element_count"] == 3, "mixed surface does not exclude volume from check");
        b.Geometry.Items.RemoveAt(1);
        b.Geometry.Items.Add(new Solid { Volume = 0 });
        b.Geometry.Items.Add(new Face());
        result = Map(InterferenceService.Check(doc, request));
        details = Map(Map(Items(result["geometry_notes"])[0])["geometry_details"]);
        Assert((int)details["zero_volume_solid_count"] == 1 && (int)details["face_count"] == 1 && (int)details["mesh_count"] == 0, "zero volume and faces distinguished from mesh");
        b.Geometry.Items.RemoveRange(1, 2);
        var solidGeometry = b.Geometry;
        b.Geometry = new GeometryElement { Items = new List<GeometryObject> { new Mesh() } };
        Assert((bool)Map(InterferenceService.Check(doc, request))["complete"], "mesh-only element excluded by test model policy");
        b.Geometry = solidGeometry;
        b.ThrowGeometry = true;
        Assert(Items(Map(InterferenceService.Check(doc, request))["errors"]).Count == 1, "geometry errors retained");
        b.ThrowGeometry = false;
        var savedGeometry = b.Geometry;
        b.Geometry = null;
        b.Box = new BoundingBoxXYZ { Max = new XYZ { X = 1, Y = 1, Z = 1 } };
        result = Map(InterferenceService.Check(doc, request));
        Assert((bool)result["complete"] && Items(result["non_volumetric_elements"]).Count == 1, "bounds without volume explicitly excluded by test model policy");
        b.Geometry = savedGeometry;
        b.Box = null;
        var link = new RevitLinkInstance { Id = new ElementId(8), UniqueId = "link" };
        doc.Elements.Add(link);
        Assert(!(bool)Map(InterferenceService.Check(doc, request))["complete"], "linked models not silently omitted");
        doc.Elements.Remove(link);
        doc.IsWorkshared = true;
        doc.Worksets.Add(new Workset { Name = "closed", IsOpen = false });
        Assert(!(bool)Map(InterferenceService.Check(doc, request))["complete"], "closed workset means incomplete");
        doc.IsWorkshared = false;
        var noVolume = new Element { Id = new ElementId(9), UniqueId = "curve", Geometry = new GeometryElement { Items = new List<GeometryObject> { new Curve() } } };
        doc.Elements.Add(noVolume);
        result = Map(InterferenceService.Check(doc, request));
        Assert(Items(result["non_volumetric_elements"]).Count == 1 && (bool)result["complete"], "curves explicitly reported outside volume checking");
        for (int i = 0; i < 2001; i++) doc.Elements.Add(new Element { Id = new ElementId(1000 + i), UniqueId = "large" + i, Category = category });
        result = Map(InterferenceService.Check(doc, request));
        Assert((int)result["solid_element_count"] == 2004 && (bool)result["complete"], "no old 2000-element category limit");


        // Экспорт: dry-run не создаёт файлов, частичные результаты не считаются полными.
        var pdfRoot = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "kpln-pdf-tests-" + Guid.NewGuid().ToString("N"));
        var pdfDoc = new Document();
        var sheetA = new ViewSheet { Id = new ElementId(101), UniqueId = "sheet-a", SheetNumber = "2" };
        var sheetB = new ViewSheet { Id = new ElementId(102), UniqueId = "sheet-b", SheetNumber = "1" };
        pdfDoc.Elements.AddRange(new Element[] { sheetA, sheetB });
        var pdfInput = Input("document_id", "doc", "expected_revision", 7);
        var preview = Map(SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot));
        Assert((bool)preview["dry_run"] && !System.IO.Directory.Exists(pdfRoot) && pdfDoc.ExportCalls == 0, "PDF dry-run has no side effects");
        Assert((string)Map(Items(preview["sheets"])[0])["sheet_number"] == "1", "PDF deterministic sheet order");
        pdfInput["dry_run"] = false;
        var exported = Map(SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot));
        Assert((bool)exported["complete"] && (int)exported["exported_count"] == 2, "all sheets exported");
        Assert(!pdfDoc.IsModified && !pdfDoc.LastOptions.HideUnreferencedViewTags, "export does not hide unreferenced tags or modify model");
        Assert(pdfDoc.LastOptions.PaperFormat == ExportPaperFormat.Default && pdfDoc.LastOptions.Combine, "native sheet-size export");
        var firstFile = Map(Items(exported["files"])[0]);
        Assert(((string)firstFile["sha256"]).Length == 64 && System.IO.File.Exists((string)firstFile["path"]), "file receipt with hash");
        var manifest = Json.Parse(System.IO.File.ReadAllText((string)exported["manifest_path"]));
        Assert((int)manifest["revision_at_request"] == 7 && (int)manifest["total_sheet_count"] == 2, "manifest provenance");
        var secondRun = Map(SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot));
        Assert((string)secondRun["output_directory"] != (string)exported["output_directory"], "repeated export never overwrites earlier PDFs");
        sheetB.IsPlaceholder = true;
        var skippedPdf = Map(SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot));
        Assert(!(bool)skippedPdf["complete"] && Items(skippedPdf["skipped"]).Count == 1 && (int)skippedPdf["exported_count"] == 1, "placeholder explicitly incomplete");
        sheetB.IsPlaceholder = false;
        foreach (var mode in new[] { "throw", "false", "missing", "invalid" })
        {
            pdfDoc.ExportModes[102] = mode;
            var failedPdf = Map(SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot));
            Assert(!(bool)failedPdf["complete"] && Items(failedPdf["failures"]).Count == 1 && (int)failedPdf["exported_count"] == 1, "PDF failure reported: " + mode);
        }
        Reject(() => SheetPdfExportService.Export(new Document(), pdfInput, pdfRoot), "no_sheets");
        pdfDoc.IsFamilyDocument = true;
        Reject(() => SheetPdfExportService.Export(pdfDoc, pdfInput, pdfRoot), "not_project_document");
        Console.WriteLine("PASS: " + _checks + " inspection contract assertions (synthetic Revit API; not a live-model run).");
    }
}
