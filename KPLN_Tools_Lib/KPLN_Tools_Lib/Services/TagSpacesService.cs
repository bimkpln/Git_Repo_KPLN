using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_Lib.Services
{
    internal sealed class TagSpacesService
    {
        internal TagSpacesResult Create(Document document, IEnumerable<ViewPlan> views,
            FamilySymbol tagType, IEnumerable<RevitLinkInstance> links, TagSpacesEquipmentOptions equipmentOptions)
        {
            RevitLinkInstance[] selectedLinks = links.ToArray();
            bool isRoom = tagType is RoomTagType;
            if (isRoom && equipmentOptions.Enabled)
                throw new InvalidOperationException("Проверка оборудования доступна только для пространств.");
            if (!isRoom && (!(tagType is SpaceTagType) || selectedLinks.Any(link => link != null)))
                throw new InvalidOperationException("Пространства маркируются только в текущей модели маркой пространства.");

            var result = new TagSpacesResult();
            ViewPlan[] selectedViews = views.ToArray();
            var sources = new List<Tuple<RevitLinkInstance, Transform, Room[]>>();
            foreach (RevitLinkInstance link in selectedLinks)
            {
                Document source = (link == null ? document : link.GetLinkDocument()) ?? throw new InvalidOperationException("Выбранная связь выгружена.");

                Transform transform = link == null ? Transform.Identity : link.GetTransform();
                if (!transform.BasisZ.IsAlmostEqualTo(XYZ.BasisZ))
                    throw new InvalidOperationException("Маркировка помещений наклонённой связи не поддерживается.");

                Room[] sourceRooms = !isRoom ? null : new FilteredElementCollector(source)
                    .OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().Cast<Room>().ToArray();
                sources.Add(Tuple.Create(link, transform, sourceRooms));
            }

            // Одна транзакция и один обработчик ошибок на весь запуск.
            var createdTags = new Dictionary<SpatialElementTag, string>();
            var existingByScope = new Dictionary<string, HashSet<string>>();
            using (var transaction = new Transaction(document, "KPLN: Маркировать помещения / пространства"))
            {
                transaction.Start();
                var failures = new TagFailures();
                FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(failures);
                options.SetClearAfterRollback(true);
                transaction.SetFailureHandlingOptions(options);

                foreach (ViewPlan view in selectedViews)
                {
                    foreach (var source in sources)
                    {
                        try
                        {
                            CreateOnView(document, view, tagType, source.Item1, source.Item2, source.Item3, equipmentOptions, result, createdTags, existingByScope);
                        }
                        catch (Autodesk.Revit.Exceptions.RegenerationFailedException)
                        {
                            throw;
                        }
                        catch (Exception ex)
                        {
                            Skip(result, "План «" + view.Name + "»", ex.Message);
                        }
                    }
                }

                if (createdTags.Count == 0)
                {
                    transaction.RollBack();
                    return result;
                }

                // Проверяем привязки всех новых марок после одной общей регенерации.
                document.Regenerate();
                foreach (var pair in createdTags.ToArray())
                {
                    SpatialElementTag tag = pair.Key;
                    if (!tag.IsValidObject || tag.IsOrphaned)
                    {
                        if (tag.IsValidObject)
                            document.Delete(tag.Id);

                        Skip(result, pair.Value, "Марка не связалась с помещением / пространством.");
                        createdTags.Remove(tag);
                    }
                }

                if (createdTags.Count == 0)
                {
                    transaction.RollBack();
                    return result;
                }

                RemoveInvisibleNewTags(document, selectedViews, isRoom, createdTags);
                if (createdTags.Count == 0)
                {
                    transaction.RollBack();
                    return result;
                }

                ReplaceHiddenOldTags(document, selectedViews, isRoom, createdTags);

                TransactionStatus status = transaction.Commit();
                result.Messages.AddRange(failures.Messages);
                if (status != TransactionStatus.Committed)
                {
                    result.Errors.Add("Revit отменил общую транзакцию маркировки. Новые марки не сохранены.");
                    return result;
                }

                result.Created = createdTags.Keys.Count(tag => tag.IsValidObject);
            }

            return result;
        }

        private static void RemoveInvisibleNewTags(Document document, ViewPlan[] views, bool isRoom,
            Dictionary<SpatialElementTag, string> createdTags)
        {
            BuiltInCategory category = isRoom ? BuiltInCategory.OST_RoomTags : BuiltInCategory.OST_MEPSpaceTags;
            var visibleByPrimaryView = new Dictionary<ElementId, HashSet<ElementId>>();

            // Сначала получаем снимки всех выбранных планов, затем выполняем удаления.
            // Общая марка основного/зависимого вида сохраняется, если видна хотя бы на одном выбранном плане.
            foreach (ViewPlan view in views)
            {
                ElementId primaryId = GetPrimaryId(view);
                if (!visibleByPrimaryView.TryGetValue(primaryId, out HashSet<ElementId> visible))
                {
                    visible = new HashSet<ElementId>();
                    visibleByPrimaryView.Add(primaryId, visible);
                }

                using (var collector = new FilteredElementCollector(document, view.Id))
                    visible.UnionWith(collector.OfCategory(category).WhereElementIsNotElementType().ToElementIds());
            }

            foreach (var pair in createdTags.ToArray())
            {
                SpatialElementTag tag = pair.Key;
                if (!tag.IsValidObject)
                {
                    createdTags.Remove(tag);
                    continue;
                }

                View owner = document.GetElement(tag.OwnerViewId) as View ?? throw new InvalidOperationException("Не определён вид новой марки.");

                ElementId primaryId = GetPrimaryId(owner);
                if (!visibleByPrimaryView.TryGetValue(primaryId, out HashSet<ElementId> visible))
                    throw new InvalidOperationException("Не определена видимость новой марки на выбранных планах.");

                if (!visible.Contains(tag.Id))
                {
                    document.Delete(tag.Id);
                    createdTags.Remove(tag);
                }
            }

        }

        private static void ReplaceHiddenOldTags(Document document, ViewPlan[] views, bool isRoom,
            Dictionary<SpatialElementTag, string> createdTags)
        {
            var newIds = new HashSet<ElementId>(createdTags.Keys.Where(tag => tag.IsValidObject).Select(tag => tag.Id));
            var deleteIds = new HashSet<ElementId>();
            BuiltInCategory category = isRoom ? BuiltInCategory.OST_RoomTags : BuiltInCategory.OST_MEPSpaceTags;
            Element[] oldTags;
            using (var collector = new FilteredElementCollector(document))
                oldTags = collector.OfCategory(category).WhereElementIsNotElementType()
                    .Where(tag => !newIds.Contains(tag.Id)).ToArray();

            foreach (ViewPlan view in views)
            {
                if (view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)?.AsInteger() != 1)
                    continue;

                HashSet<ElementId> visibleIds;
                using (var collector = new FilteredElementCollector(document, view.Id))
                    visibleIds = new HashSet<ElementId>(collector.OfCategory(category).WhereElementIsNotElementType().ToElementIds());

                // Старую марку удаляем только при наличии видимой новой марки той же цели на этом плане.
                var replacementKeys = new HashSet<string>(createdTags.Keys
                    .Where(tag => tag.IsValidObject && visibleIds.Contains(tag.Id))
                    .Select(GetTagTargetKey).Where(key => key != null));
                ElementId primaryId = GetPrimaryId(view);
                foreach (Element oldTag in oldTags)
                {
                    if (visibleIds.Contains(oldTag.Id))
                        continue;

                    View owner = document.GetElement(oldTag.OwnerViewId) as View;
                    if (owner == null || !GetPrimaryId(owner).Equals(primaryId))
                        continue;

                    string key = GetTagTargetKey(oldTag);
                    if (key != null && replacementKeys.Contains(key))
                        deleteIds.Add(oldTag.Id);
                }
            }

            // Удаление и создание находятся в одной транзакции: при отказе Revit откатываются вместе.
            if (deleteIds.Count > 0)
                document.Delete(deleteIds.ToArray());
        }

        private static string GetTagTargetKey(Element element)
        {
            if (element is RoomTag roomTag && !roomTag.IsOrphaned)
            {
                LinkElementId id = roomTag.TaggedRoomId;
                if (id != null)
                    return roomTag.IsTaggingLink ? GetKey(id.LinkInstanceId, id.LinkedElementId)
                        : GetKey(null, roomTag.TaggedLocalRoomId);
            }
            else if (element is SpaceTag spaceTag && spaceTag.Space != null)
                return GetKey(null, spaceTag.Space.Id);

            return null;
        }

        private static void CreateOnView(Document document, ViewPlan view, FamilySymbol tagType,
            RevitLinkInstance link, Transform transform, Room[] sourceRooms,
            TagSpacesEquipmentOptions equipmentOptions, TagSpacesResult result,
            Dictionary<SpatialElementTag, string> createdTags, Dictionary<string, HashSet<string>> existingByScope)
        {
            bool isRoom = tagType is RoomTagType;
            BuiltInCategory tagCategory = isRoom ? BuiltInCategory.OST_RoomTags : BuiltInCategory.OST_MEPSpaceTags;
            if (view.GetCategoryHidden(new ElementId(tagCategory)))
            {
                Skip(result, "План «" + view.Name + "»", "категория марок скрыта на виде.");
                return;
            }

            // Для всех версий — исходный поиск: помещения документа либо пространства вида.
            // Невидимые новые марки удаляются общей очисткой после регенерации.
            SpatialElement[] candidates = isRoom
                ? sourceRooms
                : new FilteredElementCollector(document, view.Id)
                    .OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType()
                    .Cast<SpatialElement>().ToArray();

            if (!isRoom && candidates.Length == 0)
            {
                result.Skipped++;
                result.Errors.Add(string.Format("План «{0}»: пространства отсутствуют либо скрыты на виде.", view.Name));
                return;
            }

            bool visibleOnly = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)?.AsInteger() == 1;
            string scopeKey = visibleOnly ? "view:" + view.Id : "primary:" + GetPrimaryId(view);
            if (!existingByScope.TryGetValue(scopeKey, out HashSet<string> existing))
            {
                existing = GetExistingTags(document, view, isRoom, visibleOnly);
                existingByScope.Add(scopeKey, existing);
            }
            var equipment = new TagSpacesEquipmentAnalysis();
            if (!isRoom && equipmentOptions.Enabled)
            {
                ElementId phaseId = view.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId();
                if (phaseId == null || phaseId.Equals(ElementId.InvalidElementId))
                    throw new InvalidOperationException("У плана не задана стадия для анализа оборудования.");

                PlanPoint[][] crop = GetCropPolygons(view);
                equipment = new TagSpacesEquipmentService().Analyze(document, view, (Phase)document.GetElement(phaseId),
                    candidates.Cast<Space>().ToArray(), crop, equipmentOptions);
                result.Messages.AddRange(equipment.Messages);
                if (!string.IsNullOrEmpty(equipment.LayoutError))
                {
                    result.ErrorViews++;
                    result.Errors.Add(equipment.LayoutError);
                    return;
                }
            }

            int processed = 0;
            foreach (SpatialElement element in candidates)
            {
                string label = string.Format("План «{0}», {1} [id: {2}]{3}", view.Name,
                    isRoom ? "помещение" : "пространство", element.Id,
                    link == null ? string.Empty : string.Format(", связь «{0}» [id: {1}]", link.Name, link.Id));

                var location = element.Location as LocationPoint;
                if (location == null || (element is Room room && room.Area <= 0)
                    || (element is Space space && space.Area <= 0))
                {
                    Skip(result, label, "элемент не размещён, не замкнут или имеет нулевую площадь.");
                    continue;
                }

                XYZ point = transform.OfPoint(location.Point);

                processed++;
                string key = GetKey(link?.Id, element.Id);
                if (existing.Contains(key))
                {
                    result.Existing++;
                    continue;
                }

                FamilySymbol selectedType = equipment.GetTagType(element.Id, tagType, equipmentOptions);
                if (CreateTag(document, view, selectedType, link, element, point, label, result, createdTags))
                    existing.Add(key);
            }

            if (processed == 0)
                result.Messages.Add(string.Format("План «{0}»: нет размещённых элементов для маркировки "
                    + "в выбранном источнике (помещения) или на выбранном плане (пространства).", view.Name));
        }

        private static bool CreateTag(Document document, ViewPlan view, FamilySymbol type, RevitLinkInstance link,
            SpatialElement element, XYZ point, string label, TagSpacesResult result,
            Dictionary<SpatialElementTag, string> createdTags)
        {
            SpatialElementTag tag = null;
            try
            {
                var uv = new UV(point.X, point.Y);
                if (type is RoomTagType roomType)
                {
                    var roomId = link == null ? new LinkElementId(element.Id) : new LinkElementId(link.Id, element.Id);
                    RoomTag roomTag = document.Create.NewRoomTag(roomId, uv, view.Id) ?? throw new InvalidOperationException("Revit не создал марку помещения.");

                    tag = roomTag;
                    roomTag.RoomTagType = roomType;
                }
                else
                {
                    SpaceTag spaceTag = document.Create.NewSpaceTag((Space)element, uv, view) ?? throw new InvalidOperationException("Revit не создал марку пространства.");

                    tag = spaceTag;
                    spaceTag.SpaceTagType = (SpaceTagType)type;
                }

                createdTags.Add(tag, label);
                return true;
            }
            catch (Autodesk.Revit.Exceptions.RegenerationFailedException)
            {
                // После ошибки регенерации документ нельзя читать: откатываем общую транзакцию.
                throw;
            }
            catch (Exception ex)
            {
                // Удаляем частично созданную марку, если назначение типа завершилось ошибкой.
                if (tag != null && tag.IsValidObject)
                    document.Delete(tag.Id);

                Skip(result, label, ex.Message);
                return false;
            }
        }

        private static HashSet<string> GetExistingTags(Document document, ViewPlan view, bool isRoom, bool visibleOnly)
        {
            var keys = new HashSet<string>();
            BuiltInCategory category = isRoom ? BuiltInCategory.OST_RoomTags : BuiltInCategory.OST_MEPSpaceTags;
            Element[] tags;
            if (visibleOnly)
            {
                // При обрезке аннотаций скрытая марка не блокирует повторную маркировку на данном плане.
                using (var collector = new FilteredElementCollector(document, view.Id))
                    tags = collector.OfCategory(category).WhereElementIsNotElementType().ToElements().ToArray();
            }
            else
            {
                // Без обрезки сохраняем прежнюю проверку, включая скрытые марки общих видов.
                ElementId primaryId = GetPrimaryId(view);
                var owners = new HashSet<ElementId>(new FilteredElementCollector(document).OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>().Where(item => !item.IsTemplate && GetPrimaryId(item).Equals(primaryId)).Select(item => item.Id));
                using (var collector = new FilteredElementCollector(document))
                    tags = collector.OfCategory(category).WhereElementIsNotElementType()
                        .Where(tag => owners.Contains(tag.OwnerViewId)).ToArray();
            }

            foreach (Element element in tags)
            {
                if (element is RoomTag roomTag && !roomTag.IsOrphaned)
                {
                    LinkElementId id = roomTag.TaggedRoomId;
                    if (id != null)
                        keys.Add(roomTag.IsTaggingLink ? GetKey(id.LinkInstanceId, id.LinkedElementId)
                            : GetKey(null, roomTag.TaggedLocalRoomId));
                }
                else if (element is SpaceTag spaceTag && spaceTag.Space != null)
                    keys.Add(GetKey(null, spaceTag.Space.Id));
            }

            return keys;
        }

        private static PlanPoint[][] GetCropPolygons(ViewPlan view)
        {
            if (!view.CropBoxActive)
                return new PlanPoint[0][];

            using (ViewCropRegionShapeManager manager = view.GetCropRegionShapeManager())
            {
                if (manager.Split)
                    throw new InvalidOperationException("Разделённая область обрезки не поддерживается: маркируйте на плане без разделения обрезки.");

                return manager.GetCropShape().Select(loop => loop.SelectMany(curve => curve.Tessellate().Skip(1))
                    .Select(point => new PlanPoint(point.DotProduct(view.RightDirection), point.DotProduct(view.UpDirection)))
                    .ToArray()).ToArray();
            }
        }

        private static ElementId GetPrimaryId(View view)
        {
            ElementId primaryId = view.GetPrimaryViewId();
            return primaryId.Equals(ElementId.InvalidElementId) ? view.Id : primaryId;
        }

        private static string GetKey(ElementId linkId, ElementId elementId)
            => (linkId ?? ElementId.InvalidElementId).ToString() + ":" + elementId;

        private static void Skip(TagSpacesResult result, string label, string reason)
        {
            result.Skipped++;
            result.Messages.Add(label + ": " + reason);
        }

        private sealed class TagFailures : IFailuresPreprocessor
        {
            internal List<string> Messages { get; } = new List<string>();

            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                bool hasErrors = false;
                foreach (FailureMessageAccessor failure in accessor.GetFailureMessages())
                {
                    Messages.Add(failure.GetDescriptionText());
                    if (failure.GetSeverity() == FailureSeverity.Warning)
                        accessor.DeleteWarning(failure);
                    else
                        hasErrors = true;
                }

                return hasErrors ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
