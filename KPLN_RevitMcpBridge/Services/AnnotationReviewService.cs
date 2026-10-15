using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class AnnotationReviewService
    {
        internal static object CloneRoomSchedule(Document doc, IDictionary<string, object> row)
        {
            var source = DocumentationReadService.Require<ViewSchedule>(doc, row, "source_schedule_unique_id");
            var view = DocumentationReadService.Require<ViewPlan>(doc, row, "view_unique_id");
            var link = DocumentationReadService.Require<RevitLinkInstance>(doc, row, "link_instance_unique_id");
            var linked = link.GetLinkDocument();
            if (linked == null || source.Definition.CategoryId != new ElementId(BuiltInCategory.OST_Rooms)
                || !source.Definition.IncludeLinkedFiles || source.Definition.HasEmbeddedSchedule)
                throw new BridgeException("unsupported_schedule", "Нужна экспликация помещений связи без вложенной спецификации.");

            var name = row.Text("name", true);
            var prefix = row.Text("number_prefix", true);
            if (prefix.Length < 3)
                throw new BridgeException("invalid_prefix", "Задайте конкретный префикс номеров этажа.");
            if (new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Any(v => v.Name == name))
                throw new BridgeException("duplicate_name", "Вид с таким именем уже существует.");

            var phaseId = view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId();
            var map = ((RevitLinkType)doc.GetElement(link.GetTypeId())).GetPhaseMap();
            if (phaseId == null || !map.ContainsKey(phaseId))
                throw new BridgeException("missing_phase", "Нет соответствия стадии вида и связи.");
            var linkedPhase = map[phaseId];
            var transform = link.GetTransform();
            var polygons = new List<XYZ[]>();
            if (!view.CropBoxActive)
                throw new BridgeException("missing_crop", "Для фрагмента нужна активная обрезка вида.");
            using (var manager = view.GetCropRegionShapeManager())
            {
                if (manager.NumberOfSplitRegions > 1)
                    throw new BridgeException("split_crop", "Разделённая обрезка не поддерживается.");
                foreach (var loop in manager.GetCropShape())
                    polygons.Add(loop.SelectMany(c => c.Tessellate()).ToArray());
            }
            if (polygons.Count == 0)
                throw new BridgeException("missing_crop", "Не прочитана граница фрагмента.");

            var rooms = new FilteredElementCollector(linked).OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Autodesk.Revit.DB.Architecture.Room>()
                .Where(r => r.Area > 1e-9 && r.Location is LocationPoint && r.get_Parameter(BuiltInParameter.ROOM_PHASE_ID)?.AsElementId() == linkedPhase
                    && r.Number.StartsWith(prefix, StringComparison.Ordinal)).ToArray();
            var included = rooms.Where(r => polygons.Any(p => Inside(transform.OfPoint(((LocationPoint)r.Location).Point), p))).ToArray();
            var excluded = rooms.Except(included).Select(r => r.Number).Distinct().ToArray();
            if (included.Length == 0 || excluded.Length > 7)
                throw new BridgeException("unsupported_scope", "Нет помещений или для точного отбора нужно больше восьми фильтров.");
            if (rooms.GroupBy(r => r.Number).Any(g => g.Count() > 1))
                throw new BridgeException("ambiguous_rooms", "Повторяются номера помещений в выбранной стадии связи.");

            var copy = (ViewSchedule)doc.GetElement(source.Duplicate(ViewDuplicateOption.Duplicate));
            copy.Name = name;
            copy.ViewTemplateId = ElementId.InvalidElementId;
            var definition = copy.Definition;
            var numberField = definition.GetFieldOrder().Select(definition.GetField)
                .Single(f => f.ParameterId == new ElementId(BuiltInParameter.ROOM_NUMBER));
            definition.ClearFilters();
            definition.AddFilter(new ScheduleFilter(numberField.FieldId, ScheduleFilterType.BeginsWith, prefix));
            foreach (var number in excluded)
                definition.AddFilter(new ScheduleFilter(numberField.FieldId, ScheduleFilterType.NotEqual, number));
            copy.get_Parameter(BuiltInParameter.VIEW_PHASE).Set(phaseId);
            doc.Regenerate();

            // Проверяем результат фильтра по номерам фактических строк таблицы.
            var visibleFields = definition.GetFieldOrder().Select(definition.GetField).Where(f => !f.IsHidden).ToArray();
            var column = Array.FindIndex(visibleFields, f => f.FieldId == numberField.FieldId);
            var body = copy.GetTableData().GetSectionData(SectionType.Body);
            var actual = new HashSet<string>();
            for (int i = body.FirstRowNumber; i <= body.LastRowNumber; i++)
            {
                var value = copy.GetCellText(SectionType.Body, i, body.FirstColumnNumber + column).Trim();
                if (value.StartsWith(prefix, StringComparison.Ordinal)) actual.Add(value);
            }
            if (!actual.SetEquals(included.Select(r => r.Number)))
                throw new BridgeException("schedule_scope_mismatch", "Строки копии не совпали с помещениями фрагмента; пакет отменён.");

            return new { element = RevitService.ElementInfo(copy),
                included_numbers = included.Select(r => r.Number).ToArray(), excluded_numbers = excluded,
                scope_method = "room_location_inside_view_crop_in_link_phase", count = included.Length };
        }

        private static bool Inside(XYZ point, XYZ[] polygon)
        {
            bool inside = false;
            for (int i = 0, j = polygon.Length - 1; i < polygon.Length; j = i++)
            {
                var a = polygon[i]; var b = polygon[j];
                if ((a.Y > point.Y) != (b.Y > point.Y)
                    && point.X < (b.X - a.X) * (point.Y - a.Y) / (b.Y - a.Y) + a.X)
                    inside = !inside;
            }
            return inside;
        }

        internal static object DeleteTag(Document doc, IDictionary<string, object> row)
        {
            var view = DocumentationReadService.Require<View>(doc, row, "view_unique_id");
            var tag = DocumentationReadService.Require<Element>(doc, row, "tag_unique_id");
            if (!(tag is IndependentTag) && !(tag is SpatialElementTag))
                throw new BridgeException("wrong_element_type", "Удалять разрешено только марки.");
            if (tag.OwnerViewId != view.Id || tag.Pinned || tag.GroupId != ElementId.InvalidElementId)
                throw new BridgeException("invalid_tag", "Марка другого вида, закреплена или входит в группу.");

            var info = RevitService.ElementInfo(tag);
            var id = tag.Id;
            var deleted = doc.Delete(id);
            if (deleted.Any(x => x != id))
                throw new BridgeException("dependent_elements", "Удаление затрагивает другие элементы; пакет отменён.");

            return new { element = info, deleted = true };
        }

        internal static object Spaces(Document doc, IDictionary<string, object> input)
        {
            var view = DocumentationReadService.Require<ViewPlan>(doc, input, "view_unique_id");
            var radiators = input.List("radiator_unique_ids", 100).Cast<object>()
                .Select(x => doc.GetElement(x as string) as FamilyInstance
                    ?? throw new BridgeException("wrong_element_type", "Нужны экземпляры радиаторов.")).ToArray();
            var phaseId = view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId();
            var phase = phaseId == null ? null : doc.GetElement(phaseId) as Phase;
            if (phase == null)
                throw new BridgeException("missing_phase", "Не определена стадия вида.");

            var spaces = new FilteredElementCollector(doc, view.Id).OfClass(typeof(SpatialElement))
                .OfType<Space>().Where(s => s.Area > 1e-9 && s.Location is LocationPoint)
                .Where(s => s.get_Parameter(BuiltInParameter.ROOM_PHASE_ID)?.AsElementId() == phaseId)
                .ToArray();

            // Читаем штатную принадлежность и точку расчёта помещения.
            // При её отсутствии сопоставление выполняется оператором по плану.
            var associations = radiators.Select(r => {
                var native = r.get_Space(phase);
                var candidates = new List<string>();
                if (native != null && spaces.Any(s => s.Id == native.Id))
                    candidates.Add(native.UniqueId);
                return new { radiator = RevitService.ElementInfo(r),
                    space_unique_ids = candidates.ToArray(),
                    calculation_point_mm = r.HasSpatialElementCalculationPoint ? Point(r.GetSpatialElementCalculationPoint()) : null,
                    method = candidates.Count == 1 ? "Revit_family_space" : "review_plan_required",
                    ambiguous = candidates.Count != 1 };
            }).ToArray();

            using (var options = new SpatialElementBoundaryOptions())
            {
                return new { view = RevitService.ElementInfo(view), phase = RevitService.ElementInfo(phase),
                    radiator_associations = associations,
                    spaces = spaces.Select(s => new { element = RevitService.ElementInfo(s),
                        number = s.Number, area_m2 = s.Area * 0.09290304,
                        location_mm = Point(((LocationPoint)s.Location).Point),
                        boundary_loops_mm = s.GetBoundarySegments(options)?.Select(loop =>
                            loop.SelectMany(segment => segment.GetCurve().Tessellate()).Select(Point).ToArray()).ToArray()
                    }).ToArray() };
            }
        }

        private static double[] Point(XYZ p) => new[] { p.X * 304.8, p.Y * 304.8, p.Z * 304.8 };
    }
}
