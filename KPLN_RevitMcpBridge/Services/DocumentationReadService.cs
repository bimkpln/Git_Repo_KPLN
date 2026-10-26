using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Факты образца оформления; проектные решения остаются в наборе правил.</summary>
    internal static class DocumentationReadService
    {
        internal static T Require<T>(Document doc, IDictionary<string, object> input, string key) where T : Element
        {
            var value = doc.GetElement(input.Text(key, true)) as T;
            if (value == null)
                throw new BridgeException("wrong_element_type", key + ": элемент отсутствует или имеет неподходящий класс.");

            return value;
        }

        internal static double[] Point(XYZ p) => new[] { p.X * 304.8, p.Y * 304.8, p.Z * 304.8 };
        private static double[] Vector(XYZ p) => new[] { p.X, p.Y, p.Z };
        internal static object Info(Element e) => e == null ? null : RevitService.ElementInfo(e);

        internal static object Layout(Document doc, IDictionary<string, object> input)
        {
            var sheet = Require<ViewSheet>(doc, input, "sheet_unique_id");
            var titles = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType().Where(e => e.OwnerViewId == sheet.Id).ToArray();
            return new
            {
                sheet = Info(sheet), number = sheet.SheetNumber, units = "mm", coordinate_system = "sheet",
                outline_min_mm = new[] { sheet.Outline.Min.U * 304.8, sheet.Outline.Min.V * 304.8 },
                outline_max_mm = new[] { sheet.Outline.Max.U * 304.8, sheet.Outline.Max.V * 304.8 },
                title_blocks = titles.Select(e => new { instance = Info(e), type = Info(doc.GetElement(e.GetTypeId())) }).ToArray(),
                viewports = sheet.GetAllViewports().Select(id => (Viewport)doc.GetElement(id)).Select(vp => new
                {
                    viewport = Info(vp), view = Info(doc.GetElement(vp.ViewId)), type = Info(doc.GetElement(vp.GetTypeId())),
                    center_mm = Point(vp.GetBoxCenter()), rotation = vp.Rotation.ToString(),
                    placement_point_mm = Point(SheetLayoutService.PlacementPoint(vp)),
                    box_min_mm = Point(vp.GetBoxOutline().MinimumPoint), box_max_mm = Point(vp.GetBoxOutline().MaximumPoint),
                    label_bounds = SheetLayoutService.LabelForLayout(vp),
#if !(Debug2020 || Revit2020)
                    label_offset_mm = Point(vp.LabelOffset), label_line_length_mm = vp.LabelLineLength * 304.8,
#endif
                    scale = ((View)doc.GetElement(vp.ViewId)).Scale,
                    template = Info(doc.GetElement(((View)doc.GetElement(vp.ViewId)).ViewTemplateId))
                }).ToArray(),
                schedules = new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance)).Cast<ScheduleSheetInstance>()
                    .Where(s => s.OwnerViewId == sheet.Id).Select(s => new
                    {
                        instance = Info(s), schedule = Info(doc.GetElement(s.ScheduleId)), point_mm = Point(s.Point),
                        placement_point_mm = Point(SheetLayoutService.PlacementPoint(s)),
                        is_titleblock_revision_schedule = s.IsTitleblockRevisionSchedule,
                        bounds = SheetLayoutService.BoundsForLayout(s, sheet)
                    }).ToArray()
            };
        }

        internal static object Catalog(Document doc, IDictionary<string, object> input)
        {
            var kind = input.Text("kind", true);
            using (var collector = new FilteredElementCollector(doc))
            {
                IEnumerable<Element> values;
                switch (kind)
                {
                    case "title_blocks":
                        values = collector.OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
                            .Cast<FamilySymbol>().Where(s => s.FamilyName.StartsWith("020_", StringComparison.OrdinalIgnoreCase));
                        break;
                    case "tag_types":
                        values = collector.WhereElementIsElementType().Where(s => s.Category != null && s.Category.IsTagCategory);
                        break;
                    case "dimension_types": values = collector.OfClass(typeof(DimensionType)); break;
                    case "view_types": values = collector.OfClass(typeof(ViewFamilyType)); break;
                    case "view_templates": values = collector.OfClass(typeof(View)).Cast<View>().Where(v => v.IsTemplate); break;
                    default: throw new BridgeException("invalid_input", "Неизвестный kind каталога оформления.");
                }

                int offset = input.Range("offset", 0, 0, 10000000);
                int limit = input.Range("limit", 100, 1, 200);
                var page = values.OrderBy(e => RevitService.IdValue(e.Id)).Skip(offset).Take(limit + 1).ToArray();
                return new
                {
                    kind, offset, items = page.Take(limit).Select(e => new
                    {
                        element = Info(e), family_name = (e as FamilySymbol)?.FamilyName ?? e.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)?.AsString(),
                        view_family = (e as ViewFamilyType)?.ViewFamily.ToString(),
                        dimension_style = (e as DimensionType)?.StyleType.ToString(),
                        view_type = (e as View)?.ViewType.ToString()
                    }).ToArray(),
                    next_offset = page.Length > limit ? (int?)(offset + limit) : null
                };
            }
        }

        internal static object Levels(Document doc, IDictionary<string, object> input)
        {
            var plans = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().Where(v => !v.IsTemplate).ToArray();
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(l => l.ProjectElevation).ToArray();
            int offset = input.Range("offset", 0, 0, 10000000);
            int limit = input.Range("limit", 100, 1, 200);

            return new
            {
                selection_policy = "evidence_only; building_story is a candidate, confirm project floors and typical groups from reference/operator",
                total_count = levels.Length, offset,
                items = levels.Skip(offset).Take(limit).Select(l => new
                {
                    level = Info(l), elevation_mm = l.Elevation * 304.8, project_elevation_mm = l.ProjectElevation * 304.8,
                    building_story = l.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY)?.AsInteger() == 1,
                    structural = l.get_Parameter(BuiltInParameter.LEVEL_IS_STRUCTURAL)?.AsInteger() == 1,
                    same_elevation_level_ids = levels.Where(other => other.Id != l.Id && Math.Abs(other.ProjectElevation - l.ProjectElevation) < 1 / 304.8).Select(other => other.UniqueId).ToArray(),
                    plans = plans.Where(v => v.GenLevel != null && v.GenLevel.Id == l.Id).Select(v => new
                    {
                        view = Info(v), scale = v.Scale, template = Info(doc.GetElement(v.ViewTemplateId)),
                        sheet_number = v.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER)?.AsString()
                    }).ToArray()
                }).ToArray(),
                next_offset = offset + limit < levels.Length ? (int?)(offset + limit) : null
            };
        }

        internal static object Annotations(Document doc, IDictionary<string, object> input)
        {
            var view = Require<View>(doc, input, "view_unique_id");
            int offset = input.Range("offset", 0, 0, 10000000);
            int limit = input.Range("limit", 100, 1, 200);
            var values = new FilteredElementCollector(doc).OwnedByView(view.Id).WhereElementIsNotElementType()
                .Where(e => e is IndependentTag || e is SpatialElementTag || e is Dimension || AnnotationContentService.Supports(e))
                .OrderBy(e => RevitService.IdValue(e.Id)).Skip(offset).Take(limit + 1).ToArray();

            return new
            {
                view = Info(view), scale = view.Scale, template = Info(doc.GetElement(view.ViewTemplateId)),
                coordinate_system = "document_internal", units = "mm", origin_mm = Point(view.Origin),
                right = Vector(view.RightDirection), up = Vector(view.UpDirection), direction = Vector(view.ViewDirection),
                visibility_semantics = "owner_view_elements; crop, visibility and occlusion not evaluated",
                items = values.Take(limit).Select(e => ReadAnnotation(doc, e)).ToArray(), offset,
                next_offset = values.Length > limit ? (int?)(offset + limit) : null
            };
        }

        private static object ReadAnnotation(Document doc, Element e)
        {
            try
            {
                return Annotation(doc, e);
            }
            catch (Exception ex)
            {
                // Неподдерживаемая аннотация не должна скрывать остальные элементы страницы.
                return new { kind = "read_error", element = Info(e), error_type = ex.GetType().FullName, message = ex.Message };
            }
        }

        private static object Annotation(Document doc, Element e)
        {
            if (AnnotationContentService.Supports(e))
                return AnnotationContentService.Read(doc, e);

            var tag = e as IndependentTag;
            if (tag != null)
            {
#if Debug2020 || Revit2020
                var ids = new[] { tag.TaggedElementId };
#else
                var ids = tag.GetTaggedElementIds();
#endif
                return new
                {
                    kind = "independent_tag", element = Info(e), type = Info(doc.GetElement(e.GetTypeId())),
                    head_mm = Point(tag.TagHeadPosition), has_leader = tag.HasLeader, orientation = tag.TagOrientation.ToString(),
                    text = tag.TagText, is_orphaned = tag.IsOrphaned,
                    leader_geometry = LeaderGeometry(tag),
                    targets = ids.Select(id => Target(doc, id)).ToArray()
                };
            }

            var spatial = e as SpatialElementTag;
            if (spatial != null)
            {
                var roomTag = spatial as RoomTag;
                var spaceTag = spatial as SpaceTag;
                return new
                {
                    kind = roomTag != null ? "room_tag" : "space_tag", element = Info(e), type = Info(doc.GetElement(e.GetTypeId())),
                    head_mm = Point(spatial.TagHeadPosition), has_leader = spatial.HasLeader,
                    target = roomTag != null ? Target(doc, roomTag.TaggedRoomId) : Info(spaceTag?.Space)
                };
            }

            var dimension = (Dimension)e;
            return new
            {
                kind = "dimension", element = Info(e), type = Info(doc.GetElement(e.GetTypeId())),
                curve = CurveInfo(dimension.Curve),
                references = dimension.References.Cast<Reference>().Select(r => ReferenceInfo(doc, r)).ToArray(),
                segment_count = dimension.NumberOfSegments,
                value_override = dimension.NumberOfSegments == 0 ? dimension.ValueOverride : null,
                segments = dimension.NumberOfSegments == 0 ? new object[0] : dimension.Segments.Cast<DimensionSegment>()
                    .Select((s, index) => (object)new { index, value_override = s.ValueOverride, above = s.Above, below = s.Below, prefix = s.Prefix, suffix = s.Suffix }).ToArray()
            };
        }

        private static object LeaderGeometry(IndependentTag tag)
        {
            if (!tag.HasLeader) return null;

            try
            {
                var free = tag.LeaderEndCondition == LeaderEndCondition.Free;
#if Debug2020 || Revit2020
                return new { status = "available", condition = tag.LeaderEndCondition.ToString(),
                    end_mm = free ? Point(tag.LeaderEnd) : null,
                    elbow_mm = tag.HasElbow ? Point(tag.LeaderElbow) : null };
#else
                var references = tag.GetTaggedReferences();
                return new { status = "available", condition = tag.LeaderEndCondition.ToString(),
                    leaders = references.Select(reference => new {
                        end_mm = free ? Point(tag.GetLeaderEnd(reference)) : null,
                        elbow_mm = tag.HasLeaderElbow(reference) ? Point(tag.GetLeaderElbow(reference)) : null
                    }).ToArray() };
#endif
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                return new { status = "unavailable" };
            }
        }

        private static object Target(Document doc, LinkElementId id)
        {
            var link = doc.GetElement(id.LinkInstanceId) as RevitLinkInstance;
            return new
            {
                link = Info(link),
                element = Info(link == null ? doc.GetElement(id.HostElementId) : link.GetLinkDocument()?.GetElement(id.LinkedElementId))
            };
        }

        internal static object CurveInfo(Curve curve)
        {
            if (curve == null) return null;
            return new { kind = curve.GetType().Name, bound = curve.IsBound,
                start_mm = curve.IsBound ? Point(curve.GetEndPoint(0)) : null,
                end_mm = curve.IsBound ? Point(curve.GetEndPoint(1)) : null,
                origin_mm = curve is Line ? Point(((Line)curve).Origin) : null,
                direction = curve is Line ? Vector(((Line)curve).Direction) : null };
        }

        private static object ReferenceInfo(Document doc, Reference reference)
        {
            return new { stable_reference = reference.ConvertToStableRepresentation(doc),
                reference_type = reference.ElementReferenceType.ToString(), element = Info(doc.GetElement(reference.ElementId)) };
        }

        internal static object References(Document doc, IDictionary<string, object> input)
        {
            var element = Require<Element>(doc, input, "element_unique_id");
            var view = Require<View>(doc, input, "view_unique_id");
            var rows = new List<object>();
            var seen = new HashSet<string>();
            int max = input.Range("limit", 200, 1, 1000);
            bool truncated = false;
            Action<Reference, string, object> add = (reference, kind, geometry) =>
            {
                if (reference == null) return;
                var stable = reference.ConvertToStableRepresentation(doc);
                if (!seen.Add(stable)) return;
                if (rows.Count >= max) { truncated = true; return; }
                rows.Add(new { kind, stable_reference = stable, geometry });
            };

            if (element is Grid || element is Level || element is ReferencePlane)
                add(new Reference(element), "datum", element is Grid ? CurveInfo(((Grid)element).Curve) : null);

            var family = element as FamilyInstance;
            if (family != null)
                foreach (FamilyInstanceReferenceType kind in Enum.GetValues(typeof(FamilyInstanceReferenceType)))
                    if (kind != FamilyInstanceReferenceType.NotAReference)
                        foreach (var reference in family.GetReferences(kind))
                            add(reference, "family_" + kind, null);

            // Не применять GetInstanceGeometry: его копии не пригодны для создания размеров.
            Action<GeometryElement, Transform, int> walk = null;
            walk = (geometry, transform, depth) =>
            {
                if (geometry == null || depth > 8) return;
                foreach (var item in geometry)
                {
                    var instance = item as GeometryInstance;
                    if (instance != null)
                    {
                        walk(instance.GetSymbolGeometry(), transform.Multiply(instance.Transform), depth + 1);
                        continue;
                    }

                    var solid = item as Solid;
                    if (solid != null)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            var plane = face as PlanarFace;
                            if (plane != null)
                                add(face.Reference, "planar_face", new { origin_mm = Point(transform.OfPoint(plane.Origin)), normal = Vector(transform.OfVector(plane.FaceNormal)) });
                        }
                        foreach (Edge edge in solid.Edges)
                            add(edge.Reference, "edge", CurveInfo(edge.AsCurve().CreateTransformed(transform)));
                    }

                    var curve = item as Curve;
                    if (curve != null)
                        add(curve.Reference, "curve", CurveInfo(curve.CreateTransformed(transform)));
                }
            };

            walk(element.get_Geometry(new Options { ComputeReferences = true, IncludeNonVisibleObjects = true, View = view }), Transform.Identity, 0);
            return new { element = Info(element), view = Info(view), units = "mm", coordinate_system = "document_internal",
                items = rows.ToArray(), truncated, note = "Candidates only: Revit validates dimension compatibility in dry_run. Linked geometry is not traversed." };
        }
    }
}
