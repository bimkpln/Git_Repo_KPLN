using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Размещение и консервативная проверка габаритов в координатах листа.</summary>
    internal static class SheetLayoutService
    {
        internal sealed class Region
        {
            public string element_unique_id;
            public string kind;
            public double[] min_mm;
            public double[] max_mm;
        }

        internal sealed class Report
        {
            public string sheet_unique_id;
            public string units = "mm";
            public string coordinate_system = "sheet";
            public string method = "conservative_api_bounds_including_view_annotations_and_labels";
            public bool complete;
            public bool bounds_clear;
            public bool frame_supplied;
            public bool visual_review_required = true;
            public double minimum_gap_mm;
            public Region content_region;
            public Region[] regions;
            public object[] issues;
            public object[] unreadable;
            public string note = "Габариты могут включать пустоты. Отсутствие пересечений габаритов не заменяет PDF. Без content_min/max проверены края бумаги, а не рамка и штамп.";
        }

        internal static XYZ PlacementPoint(Element element)
        {
            var viewport = element as Viewport;
            var schedule = element as ScheduleSheetInstance;
            var point = viewport != null ? viewport.GetBoxCenter() : schedule?.Point;
            if (point == null)
                throw new BridgeException("wrong_element_type", "Нужен видовой экран или экземпляр спецификации на листе.");

            return new XYZ(point.X, point.Y, 0);
        }

        internal static XYZ SheetPoint(IDictionary<string, object> input, string key)
        {
            var point = DocumentationWriteService.Point(input, key);
            if (Math.Abs(point.Z) > 1e-9)
                throw new BridgeException("invalid_point", key + ": для листа требуется Z=0.");

            return point;
        }

        internal static Element Move(Document doc, IDictionary<string, object> row)
        {
            var sheet = DocumentationReadService.Require<ViewSheet>(doc, row, "sheet_unique_id");
            var element = DocumentationReadService.Require<Element>(doc, row, "instance_unique_id");
            var viewport = element as Viewport;
            var schedule = element as ScheduleSheetInstance;
            if (sheet.IsPlaceholder || (viewport == null && schedule == null))
                throw new BridgeException("wrong_element_type", "Перемещать можно только размещённый вид или спецификацию.");
            if (element.OwnerViewId != sheet.Id)
                throw new BridgeException("wrong_sheet", "Элемент не принадлежит указанному листу.", 409);
            if (element.Pinned || element.GroupId != ElementId.InvalidElementId)
                throw new BridgeException("protected_element", "Закреплённые и групповые элементы не перемещаются.", 409);
            if (schedule != null && schedule.IsTitleblockRevisionSchedule)
                throw new BridgeException("protected_element", "Ведомость изменений основной надписи не перемещается отдельно.", 409);

            var expected = SheetPoint(row, "expected_point_mm");
            var destination = SheetPoint(row, "point_mm");
            if (PlacementPoint(element).DistanceTo(expected) * 304.8 > 0.1)
                throw new BridgeException("stale_position", "Положение изменилось. Перечитайте get_sheet_layout.", 409);

            if (viewport != null)
                viewport.SetBoxCenter(destination);
            else
                schedule.Point = destination;

            doc.Regenerate();
            if (PlacementPoint(element).DistanceTo(destination) * 304.8 > 0.1)
                throw new BridgeException("placement_failed", "Revit не подтвердил заданное положение на листе.", 409);

            return element;
        }

        internal static Element UpdateLabel(Document doc, IDictionary<string, object> row)
        {
#if Debug2020 || Revit2020
            throw new BridgeException("unsupported_version", "Положение заголовка через API доступно начиная с Revit 2022.");
#else
            var sheet = DocumentationReadService.Require<ViewSheet>(doc, row, "sheet_unique_id");
            var viewport = DocumentationReadService.Require<Viewport>(doc, row, "viewport_unique_id");
            if (sheet.IsPlaceholder || viewport.OwnerViewId != sheet.Id)
                throw new BridgeException("wrong_sheet", "Видовой экран не принадлежит указанному листу.", 409);
            if (viewport.Pinned || viewport.GroupId != ElementId.InvalidElementId)
                throw new BridgeException("protected_element", "Закреплённые и групповые видовые экраны не изменяются.", 409);

            var expected = SheetPoint(row, "expected_label_offset_mm");
            var destination = SheetPoint(row, "label_offset_mm");
            if (viewport.LabelOffset.DistanceTo(expected) * 304.8 > 0.1)
                throw new BridgeException("stale_position", "Заголовок перемещён. Перечитайте get_sheet_layout.", 409);

            bool changeLength = row.Get("label_line_length_mm") != null;
            if (changeLength != (row.Get("expected_label_line_length_mm") != null))
                throw new BridgeException("invalid_input", "Для длины линии нужны новое и ожидаемое значения.");

            double length = viewport.LabelLineLength;
            if (changeLength)
            {
                var expectedLength = Json.Double(row.Get("expected_label_line_length_mm")) / 304.8;
                length = Json.Double(row.Get("label_line_length_mm")) / 304.8;
                if (Math.Abs(viewport.LabelLineLength - expectedLength) * 304.8 > 0.1)
                    throw new BridgeException("stale_position", "Длина линии изменена. Перечитайте get_sheet_layout.", 409);
            }


            // LabelOffset — смещение API от нижнего левого угла окна без поворота,
            // а не абсолютная точка текста. Положение самого плана сохраняется.
            var center = viewport.GetBoxCenter();
            var rotation = viewport.Rotation;
            viewport.LabelOffset = destination;
            if (changeLength)
                viewport.LabelLineLength = length;

            doc.Regenerate();
            var actualCenter = viewport.GetBoxCenter();
            if (viewport.LabelOffset.DistanceTo(destination) * 304.8 > 0.1
                || Math.Abs(viewport.LabelLineLength - length) * 304.8 > 0.1
                || Math.Abs(actualCenter.X - center.X) * 304.8 > 0.1
                || Math.Abs(actualCenter.Y - center.Y) * 304.8 > 0.1
                || viewport.Rotation != rotation)
                throw new BridgeException("placement_failed", "Revit не подтвердил положение заголовка без перемещения плана.", 409);

            return viewport;
#endif
        }

        private static Region Rectangle(string id, string kind, IEnumerable<XYZ> points)
        {
            var array = points.ToArray();
            var min = new[] { array.Min(p => p.X) * 304.8, array.Min(p => p.Y) * 304.8, 0 };
            var max = new[] { array.Max(p => p.X) * 304.8, array.Max(p => p.Y) * 304.8, 0 };
            if (min.Concat(max).Any(v => double.IsNaN(v) || double.IsInfinity(v)) || max[0] < min[0] || max[1] < min[1])
                throw new BridgeException("invalid_bounds", "Не удалось прочитать конечный габарит листа.");

            return new Region { element_unique_id = id, kind = kind, min_mm = min, max_mm = max };
        }

        internal static Region BoundingRegion(Element element, ViewSheet sheet)
        {
            var box = element.get_BoundingBox(sheet);
            if (box == null)
                throw new BridgeException("missing_bounds", "Revit не вернул габарит экземпляра на листе.");

            var region = Rectangle(element.UniqueId, "schedule", from x in new[] { box.Min.X, box.Max.X }
                from y in new[] { box.Min.Y, box.Max.Y } from z in new[] { box.Min.Z, box.Max.Z }
                select box.Transform.OfPoint(new XYZ(x, y, z)));
            if (region.max_mm[0] <= region.min_mm[0] || region.max_mm[1] <= region.min_mm[1])
                throw new BridgeException("missing_bounds", "Revit вернул пустой габарит спецификации.");

            return region;
        }

        private static Region OutlineRegion(string id, string kind, Outline outline)
        {
            if (outline == null)
                throw new BridgeException("missing_bounds", "Revit не вернул габарит вида или его заголовка.");

            return Rectangle(id, kind, new[] { outline.MinimumPoint, outline.MaximumPoint });
        }

        internal static object BoundsForLayout(Element element, ViewSheet sheet)
        {
            try { return BoundingRegion(element, sheet); }
            catch (Exception ex) { return new { status = "unreadable", message = ex.Message }; }
        }

        internal static object LabelForLayout(Viewport viewport)
        {
            try { return OutlineRegion(viewport.UniqueId, "viewport_label", viewport.GetLabelOutline()); }
            catch (Exception ex) { return new { status = "unreadable", message = ex.Message }; }
        }

        internal static bool TooClose(Region a, Region b, double gap)
        {
            return a.min_mm[0] < b.max_mm[0] + gap && b.min_mm[0] < a.max_mm[0] + gap
                && a.min_mm[1] < b.max_mm[1] + gap && b.min_mm[1] < a.max_mm[1] + gap;
        }

        internal static bool Inside(Region item, Region field)
        {
            const double tolerance = 0.01;
            return item.min_mm[0] >= field.min_mm[0] - tolerance && item.min_mm[1] >= field.min_mm[1] - tolerance
                && item.max_mm[0] <= field.max_mm[0] + tolerance && item.max_mm[1] <= field.max_mm[1] + tolerance;
        }

        private static Region InputRegion(IDictionary<string, object> input, string minKey, string maxKey, string id, string kind)
        {
            var min = SheetPoint(input, minKey);
            var max = SheetPoint(input, maxKey);
            if (max.X <= min.X || max.Y <= min.Y)
                throw new BridgeException("invalid_bounds", "Минимум поля должен быть меньше максимума по X и Y.");

            return Rectangle(id, kind, new[] { min, max });
        }

        internal static Report Check(Document doc, IDictionary<string, object> input)
        {
            var sheet = DocumentationReadService.Require<ViewSheet>(doc, input, "sheet_unique_id");
            var regions = new List<Region>();
            var unreadable = new List<object>();
            var issues = new List<object>();
            double gap = input.Get("minimum_gap_mm") == null ? 0 : Json.Double(input.Get("minimum_gap_mm"));
            if (gap < 0 || gap > 100)
                throw new BridgeException("invalid_input", "minimum_gap_mm должен быть в пределах 0–100 мм.");

            bool explicitField = input.Get("content_min_mm") != null || input.Get("content_max_mm") != null;
            var paper = Rectangle(sheet.UniqueId, "sheet_outline", new[] {
                new XYZ(sheet.Outline.Min.U, sheet.Outline.Min.V, 0), new XYZ(sheet.Outline.Max.U, sheet.Outline.Max.V, 0) });
            var field = explicitField ? InputRegion(input, "content_min_mm", "content_max_mm", sheet.UniqueId, "content_field") : paper;
            if (!Inside(field, paper))
                throw new BridgeException("invalid_bounds", "Полезное поле выходит за границы бумаги.");

            // GetBoxOutline учитывает отображаемую аннотацию; CropBox для этой проверки непригоден.
            foreach (var id in sheet.GetAllViewports())
            {
                var viewport = (Viewport)doc.GetElement(id);
                try
                {
                    regions.Add(OutlineRegion(viewport.UniqueId, "viewport", viewport.GetBoxOutline()));
                    var showLabel = doc.GetElement(viewport.GetTypeId())?.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_SHOW_LABEL);
                    if (showLabel == null || showLabel.AsInteger() != 0)
                    {
                        var label = OutlineRegion(viewport.UniqueId, "viewport_label", viewport.GetLabelOutline());
                        if (label.max_mm[0] > label.min_mm[0] || label.max_mm[1] > label.min_mm[1])
                            regions.Add(label);
                    }
                }
                catch (Exception ex) { unreadable.Add(new { element_unique_id = viewport.UniqueId, message = ex.Message }); }
            }

            foreach (var schedule in new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>().Where(s => s.OwnerViewId == sheet.Id && !s.IsTitleblockRevisionSchedule))
            {
                try { regions.Add(BoundingRegion(schedule, sheet)); }
                catch (Exception ex) { unreadable.Add(new { element_unique_id = schedule.UniqueId, message = ex.Message }); }
            }

            foreach (var region in regions.Where(r => !Inside(r, field)))
                issues.Add(new { kind = "outside_content_field", element_unique_id = region.element_unique_id, region });

            var reserved = input.Get("reserved_regions") as IList;
            if (input.Get("reserved_regions") != null && (reserved == null || reserved.Count > 100))
                throw new BridgeException("invalid_input", "reserved_regions: ожидается до 100 прямоугольных служебных зон.");
            if (reserved != null)
                foreach (var raw in reserved)
                {
                    var row = raw as IDictionary<string, object>;
                    if (row == null)
                        throw new BridgeException("invalid_input", "Служебная зона должна быть объектом.");
                    var region = InputRegion(row, "min_mm", "max_mm", "reserved:" + row.Text("name", true), "reserved");
                    foreach (var item in regions.Where(r => TooClose(r, region, gap)))
                        issues.Add(new { kind = "reserved_region_overlap", first = item, second = region });
                }

            for (int i = 0; i < regions.Count; i++)
                for (int j = i + 1; j < regions.Count; j++)
                    if (regions[i].element_unique_id != regions[j].element_unique_id && TooClose(regions[i], regions[j], gap))
                        issues.Add(new { kind = "overlap_or_insufficient_gap", first = regions[i], second = regions[j] });

            return new Report { sheet_unique_id = sheet.UniqueId, complete = unreadable.Count == 0,
                bounds_clear = unreadable.Count == 0 && issues.Count == 0, frame_supplied = explicitField,
                minimum_gap_mm = gap, content_region = field, regions = regions.ToArray(),
                issues = issues.ToArray(), unreadable = unreadable.ToArray() };
        }

        internal static void ValidateSheets(Document doc, IEnumerable<string> sheetIds)
        {
            foreach (var sheetId in sheetIds.Distinct())
            {
                var report = Check(doc, new Dictionary<string, object> { { "sheet_unique_id", sheetId } });
                if (!report.bounds_clear)
                    throw new BridgeException("sheet_layout_conflict", "Размещение откатилось: пересечение габаритов или выход за лист. "
                        + "Перекомпонуйте лист; при недостатке места увеличьте формат основной надписи. " + Json.Serialize(report), 409);
            }
        }
    }
}
