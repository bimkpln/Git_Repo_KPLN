using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Пакетная маркировка помещений связи в транзакции DocumentationWriteService.</summary>
    internal static class LinkedRoomTagService
    {
        private const double LevelTolerance = 10.0 / 304.8;
        private const int MaxTags = 1000;

        internal static object Create(Document doc, IDictionary<string, object> row, ref int totalCreated)
        {
            var view = DocumentationReadService.Require<ViewPlan>(doc, row, "view_unique_id");
            var link = DocumentationReadService.Require<RevitLinkInstance>(doc, row, "link_instance_unique_id");
            var type = DocumentationReadService.Require<RoomTagType>(doc, row, "tag_type_unique_id");
            if (view.IsTemplate || view.GenLevel == null ||
                (view.ViewType != ViewType.FloorPlan && view.ViewType != ViewType.CeilingPlan))
                throw new BridgeException("unsupported_view", "Выберите план этажа или потолка, не шаблон.");

            var linkedDoc = link.GetLinkDocument();
            if (linkedDoc == null)
                throw new BridgeException("link_unloaded", "Связь не загружена.");

            if (link.IsHidden(view) || (link.Category != null && view.GetCategoryHidden(link.Category.Id)) ||
                view.GetCategoryHidden(new ElementId(BuiltInCategory.OST_RoomTags)))
                throw new BridgeException("hidden_annotation_source", "Связь или категория марок помещений скрыта на плане.");

            string orientation = row.Text("orientation") ?? "Horizontal";
            if (orientation != "Horizontal" && orientation != "Vertical")
                throw new BridgeException("invalid_input", "orientation: Horizontal или Vertical.");


            // Координаты и отметки переводим через конкретный экземпляр связи.
            var transform = link.GetTransform();
            if (!transform.BasisZ.IsAlmostEqualTo(XYZ.BasisZ))
                throw new BridgeException("unsupported_link_transform", "Поддерживается связь с вертикальной осью Z вверх.");

            var level = ResolveLevel(linkedDoc, view, transform, row);
            var phase = ResolvePhase(doc, linkedDoc, view, link);
            var crop = GetCrop(view);
            var inverseCrop = crop == null ? null : crop.Transform.Inverse;

            // Проверяем владельца марки и оба ID: разные экземпляры одной связи независимы.
            var tagged = new HashSet<long>(new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsNotElementType().OfType<RoomTag>()
                .Where(t => t.OwnerViewId == view.Id && !t.IsOrphaned && t.TaggedRoomId.LinkInstanceId == link.Id)
                .Select(t => RevitService.IdValue(t.TaggedRoomId.LinkedElementId)));

            var rooms = new FilteredElementCollector(linkedDoc).OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType().OfType<Room>()
                .Where(r => r.LevelId == level.Id && r.get_Parameter(BuiltInParameter.ROOM_PHASE_ID)?.AsElementId() == phase.Id)
                .OrderBy(r => RevitService.IdValue(r.Id)).ToArray();
            var created = new List<object>();
            var skipped = new Dictionary<string, int>
            {
                { "already_tagged", 0 }, { "unplaced_or_unenclosed", 0 }, { "outside_crop", 0 }
            };


            foreach (var room in rooms)
            {
                var location = room.Location as LocationPoint;
                if (location == null || room.Area <= 0)
                {
                    skipped["unplaced_or_unenclosed"]++;
                    continue;
                }

                if (tagged.Contains(RevitService.IdValue(room.Id)))
                {
                    skipped["already_tagged"]++;
                    continue;
                }

                var point = transform.OfPoint(location.Point);
                if (crop != null && !InsideCrop(inverseCrop.OfPoint(point), crop))
                {
                    skipped["outside_crop"]++;
                    continue;
                }

                if (totalCreated >= MaxTags)
                    throw new BridgeException("too_many_tags", "В пакете больше 1000 новых марок. Разделите планы на несколько запросов; весь пакет отменён.");

                var tag = doc.Create.NewRoomTag(new LinkElementId(link.Id, room.Id), new UV(point.X, point.Y), view.Id);
                if (tag == null)
                    throw new BridgeException("tag_creation_failed", "Revit не создал марку помещения " + room.UniqueId);

                tag.RoomTagType = type;
                tag.HasLeader = false;
                tag.TagOrientation = orientation == "Horizontal"
                    ? SpatialElementTagOrientation.Horizontal : SpatialElementTagOrientation.Vertical;
                doc.Regenerate();

                if (tag.IsOrphaned || tag.TaggedRoomId.LinkInstanceId != link.Id || tag.TaggedRoomId.LinkedElementId != room.Id)
                    throw new BridgeException("tag_creation_failed", "Не подтверждена связь марки с помещением " + room.UniqueId);

                tagged.Add(RevitService.IdValue(room.Id));
                totalCreated++;
                created.Add(new { room_unique_id = room.UniqueId, room_number = room.Number,
                    element = RevitService.ElementInfo(tag), head_mm = DocumentationReadService.Point(tag.TagHeadPosition) });
            }

            return new
            {
                view_unique_id = view.UniqueId, link_instance_unique_id = link.UniqueId,
                linked_level = RevitService.ElementInfo(level), linked_phase = RevitService.ElementInfo(phase),
                matched_room_count = rooms.Length, created_count = created.Count, skipped,
                tags = created.ToArray(), selection = "linked_level_phase_and_location_inside_rectangular_crop"
            };
        }

        private static Level ResolveLevel(Document linkedDoc, ViewPlan view, Transform transform, IDictionary<string, object> row)
        {
            // Явное соответствие нужно, например, если в АР и ОВ разные отметки уровней.
            if (row.Text("linked_level_unique_id") != null)
                return DocumentationReadService.Require<Level>(linkedDoc, row, "linked_level_unique_id");

            var levels = new FilteredElementCollector(linkedDoc).OfClass(typeof(Level)).Cast<Level>()
                .Where(l => Math.Abs(transform.OfPoint(new XYZ(0, 0, l.ProjectElevation)).Z - view.GenLevel.ProjectElevation) <= LevelTolerance)
                .ToArray();
            if (levels.Length != 1)
                throw new BridgeException("ambiguous_linked_level", "Для плана «" + view.Name +
                    "» найдено уровней связи в пределах 10 мм: " + levels.Length + ". Укажите linked_level_unique_id явно.");

            return levels[0];
        }

        private static Phase ResolvePhase(Document doc, Document linkedDoc, ViewPlan view, RevitLinkInstance link)
        {
            var hostPhase = view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId();
            var linkType = doc.GetElement(link.GetTypeId()) as RevitLinkType;
            ElementId linkedPhase;
            if (hostPhase == null || linkType == null || !linkType.GetPhaseMap().TryGetValue(hostPhase, out linkedPhase))
                throw new BridgeException("missing_phase_mapping", "Не найдено соответствие стадии плана и стадии связи.");

            var phase = linkedDoc.GetElement(linkedPhase) as Phase;
            if (phase == null)
                throw new BridgeException("missing_phase_mapping", "Стадия связи не найдена.");

            return phase;
        }

        private static BoundingBoxXYZ GetCrop(ViewPlan view)
        {
            if (!view.CropBoxActive)
                return null;

            using (var manager = view.GetCropRegionShapeManager())
            {
                if (manager.ShapeSet || manager.NumberOfSplitRegions > 1)
                    throw new BridgeException("unsupported_crop", "Для пакетной маркировки нужна единая прямоугольная подрезка. Для сложной подрезки используйте create_revit_tags с явными целями.");
            }

            return view.CropBox ?? throw new BridgeException("invalid_crop", "Не удалось прочитать границы плана.");
        }

        private static bool InsideCrop(XYZ point, BoundingBoxXYZ crop)
        {
            const double tolerance = 0.1 / 304.8;
            return point.X >= crop.Min.X - tolerance && point.X <= crop.Max.X + tolerance &&
                point.Y >= crop.Min.Y - tolerance && point.Y <= crop.Max.Y + tolerance;
        }
    }
}
