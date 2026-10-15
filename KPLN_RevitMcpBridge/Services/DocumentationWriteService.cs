using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class DocumentationWriteService
    {
        private static T Require<T>(Document doc, IDictionary<string, object> row, string key) where T : Element
            => DocumentationReadService.Require<T>(doc, row, key);

        internal static XYZ Point(IDictionary<string, object> input, string key)
        {
            var values = input.Get(key) as IList;
            if (values == null || values.Count != 3)
                throw new BridgeException("invalid_point", key + ": нужны три конечные координаты в мм.");

            return new XYZ(Json.Double(values[0]) / 304.8, Json.Double(values[1]) / 304.8, Json.Double(values[2]) / 304.8);
        }

        private static XYZ SheetPoint(IDictionary<string, object> input, string key)
        {
            var point = Point(input, key);
            if (Math.Abs(point.Z) > 1e-9)
                throw new BridgeException("invalid_point", "Для координат листа Z должен быть 0.");

            return point;
        }

        private static XYZ UnitVector(IDictionary<string, object> input, string key)
        {
            var vector = Point(input, key) * 304.8;
            if (Math.Abs(vector.GetLength() - 1) > 1e-6)
                throw new BridgeException("invalid_basis", key + ": требуется единичный вектор.");

            return vector;
        }

        private static View View(Document doc, IDictionary<string, object> row)
        {
            var view = Require<View>(doc, row, "view_unique_id");
            if (view.IsTemplate || view is ViewSheet || view is ViewSchedule || (view is View3D && ((View3D)view).IsPerspective))
                throw new BridgeException("unsupported_view", "Нужен модельный вид для аннотаций.");

            return view;
        }

        private static string TransactionName(string command)
        {
            switch (command)
            {
                case "create_sheets": return "KPLN: создание листов из Codex";
                case "place_views": return "KPLN: размещение видов и спецификаций из Codex";
                case "move_sheet_items": return "KPLN: перемещение видов и спецификаций из Codex";
                case "update_viewport_labels": return "KPLN: положение заголовков видов из Codex";
                case "create_tags": return "KPLN: маркировка элементов из Codex";
                case "tag_linked_rooms": return "KPLN: маркировка помещений из связи";
                case "update_tags": return "KPLN: размещение марок и выносок из Codex";
                case "delete_tags": return "KPLN: удаление марок из Codex";
                case "clone_room_schedule": return "KPLN: экспликация помещений фрагмента";
                case "create_dimensions": return "KPLN: нанесение размеров из Codex";
                case "create_plan_views": return "KPLN: создание планов из Codex";
                case "create_sections": return "KPLN: создание разрезов из Codex";
                case "copy_plan_range":
                case "set_plan_range": return "KPLN: секущий диапазон плана из Codex";
                case "copy_plan_crop": return "KPLN: обрезка плана по образцу из Codex";
                default: throw new BridgeException("unknown_command", "Неизвестная команда оформления.");
            }
        }

        internal static object Execute(Document doc, Dictionary<string, object> input)
        {
            if (doc.IsReadOnly || doc.IsLinked || doc.IsModifiable || doc.IsFamilyDocument)
                throw new BridgeException("document_not_writable", "Нужен доступный для отдельной транзакции проект Revit.", 409);

            string command = input.Text("command", true);
            string transactionName = TransactionName(command);
            var rows = input.List("items", 100).Cast<object>().Select(item => item as Dictionary<string, object>
                ?? throw new BridgeException("invalid_input", "Каждая запись должна быть объектом.")).ToArray();
            bool dryRun = input.Flag("dry_run", true);
            if (command == "move_sheet_items" && rows.Select(r => r.Text("instance_unique_id", true)).Distinct().Count() != rows.Length)
                throw new BridgeException("duplicate_item", "Один экземпляр нельзя перемещать дважды в пакете.");
            var failures = new Failures();
            if (command == "update_viewport_labels" && rows.Select(r => r.Text("viewport_unique_id", true)).Distinct().Count() != rows.Length)
                throw new BridgeException("duplicate_item", "Один заголовок нельзя изменять дважды в пакете.");
            var results = new List<object>();
            int linkedRoomTagCount = 0;

            // Пробный запуск выполняет реальное создание и Commit внутри откатываемой группы.
            // При ошибке любой записи откатывается весь пакет; сохранения RVT нет.
            using (var group = new TransactionGroup(doc, transactionName))
            using (var transaction = new Transaction(doc, transactionName))
            {
                if (group.Start() != TransactionStatus.Started || transaction.Start() != TransactionStatus.Started)
                    throw new BridgeException("transaction_failed", "Не удалось начать транзакцию оформления.");

                transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions()
                    .SetFailuresPreprocessor(failures).SetClearAfterRollback(true).SetForcedModalHandling(true));
                try
                {
                    foreach (var row in rows)
                    {
                        if (command == "clone_room_schedule")
                        {
                            results.Add(AnnotationReviewService.CloneRoomSchedule(doc, row));
                            continue;
                        }
                        if (command == "delete_tags")
                        {
                            results.Add(AnnotationReviewService.DeleteTag(doc, row));
                            continue;
                        }

                        if (command == "tag_linked_rooms")
                        {
                            results.Add(LinkedRoomTagService.Create(doc, row, ref linkedRoomTagCount));
                            continue;
                        }

                        Element created;
                        switch (command)
                        {
                            case "create_sheets": created = Sheet(doc, row); break;
                            case "place_views": created = Place(doc, row); break;
                            case "move_sheet_items": created = SheetLayoutService.Move(doc, row); break;
                            case "update_viewport_labels": created = SheetLayoutService.UpdateLabel(doc, row); break;
                            case "create_tags": created = Tag(doc, row); break;
                            case "update_tags": created = UpdateTag(doc, row); break;
                            case "create_dimensions": created = Dimension(doc, row); break;
                            case "create_plan_views": created = Plan(doc, row); break;
                            case "create_sections": created = Section(doc, row); break;
                            case "copy_plan_range": created = PlanRoutingService.CopyRange(doc, row); break;
                            case "set_plan_range": created = PlanRoutingService.SetRange(doc, row); break;
                            case "copy_plan_crop": created = PlanRoutingService.CopyCrop(doc, row); break;
                            default: throw new BridgeException("unknown_command", "Неизвестная команда оформления.");
                        }

                        doc.Regenerate();
                        results.Add(new { index = results.Count, element = RevitService.ElementInfo(created) });
                    }

                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new BridgeException("transaction_rolled_back", "Пакет отменён: " + string.Join("; ", failures.Messages), 409);

                    // Проверяется итог всего пакета после Commit; при конфликте откатывается группа.
                    if (command == "place_views" || command == "move_sheet_items" || command == "update_viewport_labels")
                        SheetLayoutService.ValidateSheets(doc, rows.Select(r => r.Text("sheet_unique_id", true)));

                    var status = dryRun ? group.RollBack() : group.Assimilate();
                    if (status != (dryRun ? TransactionStatus.RolledBack : TransactionStatus.Committed))
                        throw new BridgeException("transaction_failed", "Не удалось завершить группу оформления.");
                }
                catch
                {
                    if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                    if (group.GetStatus() == TransactionStatus.Started) group.RollBack();
                    throw;
                }
            }

            return new { dry_run = dryRun, validation = "Revit_creation_and_commit", items = results.ToArray(),
                element_ids_valid = !dryRun, saved_to_disk = false,
                note = dryRun ? "Изменения откатились. ID пробного запуска недействительны. Получите свежую revision перед записью." : "Пакет применён. Модель не сохранена." };
        }

        private static Element Sheet(Document doc, IDictionary<string, object> row)
        {
            var symbol = Require<FamilySymbol>(doc, row, "title_block_type_unique_id");
            if (symbol.Category == null || RevitService.IdValue(symbol.Category.Id) != (long)BuiltInCategory.OST_TitleBlocks
                || !symbol.FamilyName.StartsWith("020_", StringComparison.OrdinalIgnoreCase))
                throw new BridgeException("invalid_title_block", "Выберите основную надпись КПЛН с кодом семейства 020_.");

            string number = row.Text("number", true);
            if (new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>().Any(s => s.SheetNumber == number))
                throw new BridgeException("duplicate_sheet_number", "Номер листа уже существует: " + number);

            if (!symbol.IsActive) symbol.Activate();
            var sheet = ViewSheet.Create(doc, symbol.Id);
            sheet.SheetNumber = number;
            sheet.Name = row.Text("name", true);
            return sheet;
        }

        private static Element Place(Document doc, IDictionary<string, object> row)
        {
            var sheet = Require<ViewSheet>(doc, row, "sheet_unique_id");
            var view = Require<View>(doc, row, "view_unique_id");
            if (sheet.IsPlaceholder || view.IsTemplate)
                throw new BridgeException("unsupported_view", "Нельзя разместить шаблон или вид на листе-заглушке.");

            // Образец задаёт центр именно окна вида, а не нижний угол PDF или рамки.
            var referenceId = row.Text("reference_viewport_unique_id");
            if ((referenceId != null) == (row.Get("point_mm") != null))
                throw new BridgeException("invalid_input", "Задайте либо point_mm, либо reference_viewport_unique_id.");

            var sample = referenceId == null ? null : Require<Viewport>(doc, row, "reference_viewport_unique_id");
            var point = sample == null ? SheetPoint(row, "point_mm") : sample.GetBoxCenter();
            var schedule = view as ViewSchedule;
            if (schedule != null)
            {
                if (sample != null)
                    throw new BridgeException("invalid_input", "Для спецификации задайте точку вставки point_mm из примера.");

                if (new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance)).Cast<ScheduleSheetInstance>()
                    .Any(s => s.OwnerViewId == sheet.Id && s.ScheduleId == schedule.Id))
                    throw new BridgeException("already_placed", "Спецификация уже размещена на листе.");

                return ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, point);
            }

            if (!Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                throw new BridgeException("cannot_place_view", "Revit не разрешает разместить вид на этом листе.");

            var viewport = Viewport.Create(doc, sheet.Id, view.Id, point);
            if (sample != null)
            {
                viewport.ChangeTypeId(sample.GetTypeId());
                viewport.Rotation = sample.Rotation;
#if !(Debug2020 || Revit2020)
                viewport.LabelOffset = sample.LabelOffset;
                viewport.LabelLineLength = sample.LabelLineLength;
#endif
            }

            doc.Regenerate();
            viewport.SetBoxCenter(point);
            // После перемещения обновляем кэш габаритов. Координаты листа двумерны:
            // служебная Z габарита окна не является координатой размещения на бумаге.
            doc.Regenerate();
            var actualCenter = viewport.GetBoxCenter();
            var planarDelta = new XYZ(actualCenter.X - point.X, actualCenter.Y - point.Y, 0);
            if (planarDelta.GetLength() > 0.1 / 304.8)
                throw new BridgeException("placement_failed", string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "Центр окна не совпал с заданным: смещение X={0:F3}, Y={1:F3} мм.",
                    planarDelta.X * 304.8, planarDelta.Y * 304.8));

            return viewport;
        }

        private static Element Tag(Document doc, IDictionary<string, object> row)
        {
            var view = View(doc, row);
            var symbol = Require<ElementType>(doc, row, "tag_type_unique_id");
            Element target;
            var head = Point(row, "head_mm");
            bool leader = row.Flag("leader");
            var kind = row.Text("kind") ?? "independent";
            var orientation = row.Text("orientation") ?? "Horizontal";
            if (orientation != "Horizontal" && orientation != "Vertical")
                throw new BridgeException("invalid_input", "orientation: Horizontal или Vertical.");
            var linkId = row.Text("link_instance_unique_id");
            RevitLinkInstance link = null;
            if (linkId != null)
            {
                // Для связей element_unique_id относится к связанному документу.
                link = Require<RevitLinkInstance>(doc, row, "link_instance_unique_id");
                target = link.GetLinkDocument()?.GetElement(row.Text("element_unique_id", true));
                if (target == null) throw new BridgeException("not_found", "Элемент связанной модели не найден.");
            }
            else
                target = Require<Element>(doc, row, "element_unique_id");

            var familySymbol = symbol as FamilySymbol;
            if (familySymbol != null && !familySymbol.IsActive) familySymbol.Activate();
            if (kind == "room" || kind == "space")
            {
                if (!(view is ViewPlan))
                    throw new BridgeException("unsupported_view", "Марка помещения/пространства требует плана.");

                SpatialElementTag spatial;
                if (kind == "room")
                {
                    if (!(target is Room)) throw new BridgeException("wrong_element_type", "Цель должна быть помещением.");
                    spatial = doc.Create.NewRoomTag(link == null ? new LinkElementId(target.Id) : new LinkElementId(link.Id, target.Id), new UV(head.X, head.Y), view.Id);
                }
                else
                {
                    if (!(target is Space) || link != null)
                        throw new BridgeException("wrong_element_type", "Поддерживаются пространства текущего документа.");
                    spatial = doc.Create.NewSpaceTag((Space)target, new UV(head.X, head.Y), view);
                }

                spatial.ChangeTypeId(symbol.Id);
                spatial.HasLeader = leader;
                spatial.TagOrientation = orientation == "Horizontal" ? SpatialElementTagOrientation.Horizontal : SpatialElementTagOrientation.Vertical;
                // NewRoomTag/NewSpaceTag уже получили UV. Их служебная глубина может
                // отличаться от уровня вида, особенно у помещений промежуточного этажа.
                doc.Regenerate();
                if (AnnotationPlane.Distance(head, spatial.TagHeadPosition, view.ViewDirection) > 0.1 / 304.8)
                    throw new BridgeException("placement_failed", "Revit не подтвердил положение пространственной марки в плоскости вида.");
                return spatial;
            }

            if (kind != "independent") throw new BridgeException("invalid_input", "Неизвестный kind марки.");
            if (familySymbol == null)
                throw new BridgeException("wrong_element_type", "Для independent нужен FamilySymbol марки.");
            var reference = new Reference(target);
            if (link != null) reference = reference.CreateLinkReference(link);
            var tag = IndependentTag.Create(doc, symbol.Id, view.Id, reference, leader,
                orientation == "Horizontal" ? TagOrientation.Horizontal : TagOrientation.Vertical, head);
            tag.TagHeadPosition = head;
            return tag;
        }

        private static Element UpdateTag(Document doc, IDictionary<string, object> row)
        {
            var view = View(doc, row);
            var element = Require<Element>(doc, row, "tag_unique_id");
            if (element.OwnerViewId != view.Id)
                throw new BridgeException("wrong_owner_view", "Марка принадлежит другому виду.");
            if (element.Pinned)
                throw new BridgeException("pinned_element", "Марка закреплена; перемещение не выполнялось.");

            var head = Point(row, "head_mm");
            var expected = Point(row, "expected_head_mm");
            var tag = element as IndependentTag;
            var spatial = element as SpatialElementTag;
            if (tag == null && spatial == null)
                throw new BridgeException("wrong_element_type", "Поддерживаются марки элементов, помещений и пространств.");

            var before = tag != null ? tag.TagHeadPosition : spatial.TagHeadPosition;
            if (before.DistanceTo(expected) > 0.1 / 304.8)
                throw new BridgeException("stale_tag_position", "Положение марки изменилось; перечитайте аннотации.", 409);
            if (Math.Abs((head - before).DotProduct(view.ViewDirection)) > 0.1 / 304.8)
                throw new BridgeException("invalid_point", "Перемещение головки должно лежать в плоскости вида.");

            bool hasElbow = row.Get("elbow_mm") != null;
            if (row.Get("leader") != null)
            {
                if (tag != null) tag.HasLeader = row.Flag("leader");
                else spatial.HasLeader = row.Flag("leader");
            }

            if (tag != null) tag.TagHeadPosition = head;
            else spatial.TagHeadPosition = head;

            if (row.Get("leader_end_mm") != null)
            {
                if (tag == null || !tag.HasLeader)
                    throw new BridgeException("invalid_leader", "Нужна независимая марка с выноской.");

                var end = Point(row, "leader_end_mm");
#if Debug2020 || Revit2020
                var target = doc.GetElement(tag.TaggedLocalElementId);
#else
                var endReferences = tag.GetTaggedReferences();
                if (endReferences.Count != 1)
                    throw new BridgeException("unsupported_leader", "Нужна одна локальная цель.");
                var target = doc.GetElement(endReferences[0].ElementId);
#endif
                var pipe = target as Autodesk.Revit.DB.Plumbing.Pipe;
                var line = (pipe?.Location as LocationCurve)?.Curve;
                var projection = line?.Project(end);
                if (pipe != null && (projection == null || projection.Distance > 2.0 / 304.8
                    || projection.Parameter < line.GetEndParameter(0) - 1e-8
                    || projection.Parameter > line.GetEndParameter(1) + 1e-8))
                    throw new BridgeException("invalid_leader_end", "Конец выноски должен лежать на оси своей трубы (допуск 2 мм).");

                // Для оборудования разрешена явная точка внутри габарита цели.
                // Это проверка принадлежности области, а не попадания на видимый контур.
                if (pipe == null)
                {
                    if (!(target is FamilyInstance) || target.Category == null
                        || !target.Category.Id.Equals(new ElementId(BuiltInCategory.OST_MechanicalEquipment)))
                        throw new BridgeException("unsupported_leader", "Поддерживаются локальные трубы и оборудование.");

                    var box = target.get_BoundingBox(null);
                    var localEnd = box?.Transform.Inverse.OfPoint(end);
                    const double tolerance = 2.0 / 304.8;
                    if (box == null || localEnd.X < box.Min.X - tolerance || localEnd.X > box.Max.X + tolerance
                        || localEnd.Y < box.Min.Y - tolerance || localEnd.Y > box.Max.Y + tolerance
                        || localEnd.Z < box.Min.Z - tolerance || localEnd.Z > box.Max.Z + tolerance)
                        throw new BridgeException("invalid_leader_end", "Конец выноски должен находиться в габарите своего оборудования (допуск 2 мм). Проверьте видимый контур в PDF.");
                }

                tag.LeaderEndCondition = LeaderEndCondition.Free;
#if Debug2020 || Revit2020
                tag.LeaderEnd = end;
#else
                tag.SetLeaderEnd(endReferences[0], end);
#endif
            }

            if (hasElbow)
            {
                var elbow = Point(row, "elbow_mm");
                if (Math.Abs((elbow - head).DotProduct(view.ViewDirection)) > 0.1 / 304.8)
                    throw new BridgeException("invalid_point", "Изгиб выноски должен лежать в плоскости головки.");
                if (!(tag != null ? tag.HasLeader : spatial.HasLeader))
                    throw new BridgeException("invalid_leader", "Для изгиба нужна включённая выноска.");

                if (tag != null)
                {
#if Debug2020 || Revit2020
                    tag.LeaderElbow = elbow;
#else
                    var references = tag.GetTaggedReferences();
                    if (references.Count != 1)
                        throw new BridgeException("unsupported_leader", "Изгиб задаётся только марке с одной целью.");
                    tag.SetLeaderElbow(references[0], elbow);
#endif
                }
                else spatial.LeaderElbow = elbow;
            }

            doc.Regenerate();
            var actual = tag != null ? tag.TagHeadPosition : spatial.TagHeadPosition;
            // Revit вправе перенести глубину головки к цели выноски. На листе важны
            // обе координаты плоскости вида; реальный сдвиг по ним остаётся ошибкой.
            if (AnnotationPlane.Distance(head, actual, view.ViewDirection) > 0.1 / 304.8)
                throw new BridgeException("placement_failed", string.Format(
                    "Головка сместилась в плоскости вида на {0:F3} мм; пакет отменён.",
                    AnnotationPlane.Distance(head, actual, view.ViewDirection) * 304.8));

            return element;
        }

        private static Element Dimension(Document doc, IDictionary<string, object> row)
        {
            var view = View(doc, row);
            var type = Require<DimensionType>(doc, row, "dimension_type_unique_id");
            if (type.StyleType != DimensionStyleType.Linear)
                throw new BridgeException("unsupported_dimension", "В этой команде поддерживаются линейные размеры и цепочки.");

            var start = Point(row, "start_mm");
            var end = Point(row, "end_mm");
            if (start.DistanceTo(end) <= doc.Application.ShortCurveTolerance)
                throw new BridgeException("invalid_dimension_line", "Размерная линия слишком короткая.");

            if (Math.Abs((end - start).Normalize().DotProduct(view.ViewDirection)) > 1e-6)
                throw new BridgeException("invalid_dimension_line", "Размерная линия должна быть параллельна плоскости вида.");

            var references = new ReferenceArray();
            var seen = new HashSet<string>();
            foreach (var value in row.List("references", 100))
            {
                var stable = value as string;
                if (string.IsNullOrWhiteSpace(stable) || !seen.Add(stable))
                    throw new BridgeException("invalid_reference", "Нужны уникальные stable_reference из Revit.");
                var reference = Reference.ParseFromStableRepresentation(doc, stable);
                if (doc.GetElement(reference.ElementId) == null)
                    throw new BridgeException("invalid_reference", "Элемент размерной привязки больше не существует.");
                references.Append(reference);
            }

            if (references.Size < 2)
                throw new BridgeException("invalid_reference", "Размеру нужны минимум две привязки.");

            var result = doc.Create.NewDimension(view, Line.CreateBound(start, end), references, type);
            if (result == null) throw new BridgeException("dimension_failed", "Revit не создал размер.");
            return result;
        }

        private static void ConfigureView(Document doc, View view, IDictionary<string, object> row)
        {
            string name = row.Text("name", true);
            if (new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>().Any(v => v.Id != view.Id && v.Name == name))
                throw new BridgeException("duplicate_view_name", "Имя вида уже существует: " + name);

            view.Name = name;
            if (row.Get("scale") != null) view.Scale = row.Range("scale", 100, 1, 24000);
            if (row.Text("template_unique_id") != null)
            {
                var template = Require<View>(doc, row, "template_unique_id");
                if (!template.IsTemplate || !view.IsValidViewTemplate(template.Id))
                    throw new BridgeException("invalid_template", "Шаблон несовместим с создаваемым видом.");
                view.ViewTemplateId = template.Id;
            }

            if (row.Get("scale") != null && view.Scale != row.Range("scale", 100, 1, 24000))
                throw new BridgeException("template_scale_conflict", "Шаблон задаёт масштаб, отличный от запрошенного.");
        }

        private static Element Plan(Document doc, IDictionary<string, object> row)
        {
            var type = Require<ViewFamilyType>(doc, row, "view_type_unique_id");
            var level = Require<Level>(doc, row, "level_unique_id");
            if (type.ViewFamily != ViewFamily.FloorPlan && type.ViewFamily != ViewFamily.CeilingPlan && type.ViewFamily != ViewFamily.StructuralPlan)
                throw new BridgeException("wrong_view_type", "Нужен тип плана этажа, потолка или несущих конструкций.");

            var plan = ViewPlan.Create(doc, type.Id, level.Id);
            ConfigureView(doc, plan, row);
            if (row.Get("crop_min_mm") != null || row.Get("crop_max_mm") != null)
            {
                var min = Point(row, "crop_min_mm");
                var max = Point(row, "crop_max_mm");
                if (min.X >= max.X || min.Y >= max.Y || min.Z >= max.Z)
                    throw new BridgeException("invalid_crop", "crop_min_mm должен быть меньше crop_max_mm по каждой оси.");
                var box = plan.CropBox;
                var inverse = box.Transform.Inverse;
                var corners = (from x in new[] { min.X, max.X } from y in new[] { min.Y, max.Y } from z in new[] { min.Z, max.Z }
                               select inverse.OfPoint(new XYZ(x, y, z))).ToArray();
                box.Min = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                box.Max = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));
                plan.CropBox = box;
                plan.CropBoxActive = true;
            }
            return plan;
        }

        private static Element Section(Document doc, IDictionary<string, object> row)
        {
            var type = Require<ViewFamilyType>(doc, row, "view_type_unique_id");
            if (type.ViewFamily != ViewFamily.Section)
                throw new BridgeException("wrong_view_type", "Нужен тип разреза.");

            var up = UnitVector(row, "up");
            var direction = UnitVector(row, "direction");
            if (Math.Abs(up.DotProduct(direction)) > 1e-6)
                throw new BridgeException("invalid_basis", "up и direction должны быть перпендикулярны.");
            var transform = Transform.Identity;
            transform.Origin = Point(row, "origin_mm");
            transform.BasisX = up.CrossProduct(direction);
            transform.BasisY = up;
            transform.BasisZ = direction;
            var min = Point(row, "min_mm");
            var max = Point(row, "max_mm");
            if (min.X >= max.X || min.Y >= max.Y || max.Z - min.Z <= doc.Application.ShortCurveTolerance)
                throw new BridgeException("invalid_section_box", "Нужен непустой объём разреза в локальных координатах его базиса.");

            var section = ViewSection.CreateSection(doc, type.Id, new BoundingBoxXYZ { Transform = transform, Min = min, Max = max });
            ConfigureView(doc, section, row);
            return section;
        }

        private sealed class Failures : IFailuresPreprocessor
        {
            internal readonly List<string> Messages = new List<string>();
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                foreach (var message in accessor.GetFailureMessages()) Messages.Add(message.GetDescriptionText());
                return Messages.Count == 0 ? FailureProcessingResult.Continue : FailureProcessingResult.ProceedWithRollBack;
            }
        }
    }
}
