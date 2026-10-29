using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Services.Schedules;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using KPLN_Library_Forms.UI.ViewModels;

namespace KPLN_ViewsAndLists_CopyByLevels.Services
{
    internal sealed class GenerationResult
    {
        internal int Created;
        internal int Failed;
        internal int Skipped;
        internal ElementId FirstSheet;
        internal string Report;
        internal string ReportPath;
    }

    internal static class GenerationService
    {
        private static readonly Guid SchemaId = new Guid("AAC7EA7E-DA3F-42C6-98B5-3AE073BC4134");

        internal static bool WasCreated(ViewSheet sheet, ViewSheet source, Level level)
        {
            Schema schema = Schema.Lookup(SchemaId);
            if (schema == null) return false;
            Entity entity = sheet.GetEntity(schema);
            return entity.IsValid() && entity.Get<string>("SourceSheet") == source.UniqueId && entity.Get<string>("Level") == level.UniqueId;
        }

        private static void Mark(ViewSheet sheet, ViewSheet source, Level level)
        {
            Schema schema = Schema.Lookup(SchemaId);
            if (schema == null)
            {
                var builder = new SchemaBuilder(SchemaId);
                builder.SetSchemaName("KPLNViewsAndListsCopyByLevelsV1");
                builder.SetReadAccessLevel(AccessLevel.Public);
                builder.SetWriteAccessLevel(AccessLevel.Public);
                builder.AddSimpleField("SourceSheet", typeof(string));
                builder.AddSimpleField("Level", typeof(string));
                schema = builder.Finish();
            }
            var entity = new Entity(schema);
            entity.Set("SourceSheet", source.UniqueId);
            entity.Set("Level", level.UniqueId);
            sheet.SetEntity(entity);
        }

        internal static GenerationResult Run(CopyByLevelsM model, ProgressInfoViewModel progress)
        {
            Document doc = model.Document;
            var result = new GenerationResult();
            int completed = 0;
            int total = model.Preview.Count;
            var report = new StringBuilder("KPLN: Листы по этажам — " + ModuleData.Version + "\r\n");
            report.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            report.AppendLine("Документ: " + doc.Title + "\r\nЭталон: " + model.SourceSheet.SheetNumber + " — " + model.SourceSheet.Name);
            report.AppendLine("Аннотации на листе: " + (model.CopySheetAnnotations ? "копировать" : "не копировать"));

            using (var group = new TransactionGroup(doc, "КПЛН: листы по этажам"))
            {
                group.Start();
                foreach (PreviewRow row in model.Preview)
                {
                    Action<string> reportStage = stage => progress.Report(completed, total,
                        "Лист " + row.Number + " — " + row.LevelName + ": " + stage);
                    reportStage("подготовка");
                    if (row.SkipExisting)
                    {
                        result.Skipped++;
                        progress.Report(++completed, total, "Пропущен существующий лист: " + row.Number);
                        report.AppendLine("\r\nПРОПУСК: " + row.LevelName + " — лист уже создан этим инструментом.");
                        continue;
                    }

                    var warnings = new List<string>();
                    using (var transaction = new Transaction(doc, "Лист " + row.Number + ": " + row.LevelName))
                    {
                        transaction.Start();
                        var failures = new GenerationFailures(warnings);
                        transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions()
                            .SetFailuresPreprocessor(failures).SetClearAfterRollback(true));
                        try
                        {
                            ViewSheet sheet = BuildSheet(model, row, warnings, reportStage);
                            Mark(sheet, model.SourceSheet, row.Level.Level);
                            ElementId createdId = sheet.Id;
                            reportStage("сохранение");
                            if (transaction.Commit() != TransactionStatus.Committed)
                                throw new InvalidOperationException("Revit откатил создание листа. " + string.Join("; ", warnings));

                            result.Created++;
                            if (result.FirstSheet == null) result.FirstSheet = createdId;
                            report.AppendLine("\r\nСОЗДАН: " + row.Number + " — " + row.Name + " [" + createdId + "]");
                        }
                        catch (Exception ex)
                        {
                            if (transaction.GetStatus() == TransactionStatus.Started) transaction.RollBack();
                            result.Failed++;
                            report.AppendLine("\r\nОТКАТ: " + row.Number + " — " + row.LevelName + "\r\n" + ex);
                        }
                        foreach (string warning in warnings.Distinct()) report.AppendLine("  • " + warning);
                    }
                    progress.Report(++completed, total, "Обработан лист: " + row.Number);
                }
                progress.Report(completed, total, "Завершение создания листов");
                if (result.Created > 0) group.Assimilate();
                else group.RollBack();
            }

            report.AppendLine("\r\nСоздано: " + result.Created + "; пропущено: " + result.Skipped + "; ошибок: " + result.Failed);
            result.Report = report.ToString();
            try
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KPLN", ModuleData.ModuleName, "Reports");
                Directory.CreateDirectory(folder);
                result.ReportPath = Path.Combine(folder, DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".txt");
                File.WriteAllText(result.ReportPath, result.Report, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                result.Report += "\r\nНе удалось сохранить файл отчёта: " + ex.Message;
            }
            return result;
        }

        private static ViewSheet BuildSheet(CopyByLevelsM model, PreviewRow row, List<string> warnings,
            Action<string> reportStage)
        {
            Document doc = model.Document;
            ViewSheet source = model.SourceSheet;
            reportStage("копирование оформления листа");
            ViewSheet sheet;
            var retainedViewports = new Dictionary<ElementId, Viewport>();
#if Debug2020 || Revit2020
            sheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
            var details = new FilteredElementCollector(doc).OwnedByView(source.Id).WhereElementIsNotElementType()
                .Where(e => !(e is Viewport) && !(e is ScheduleSheetInstance)
                    && e.GroupId == ElementId.InvalidElementId
                    && (model.CopySheetAnnotations || !IsSheetAnnotation(e))).Select(e => e.Id).ToList();
            if (details.Count > 0) ElementTransformUtils.CopyElements(source, details, sheet, Transform.Identity, new CopyPasteOptions());
            RevitHelpers.CopyParameters(source, sheet, new HashSet<ElementId>
            {
                new ElementId(BuiltInParameter.SHEET_NUMBER), new ElementId(BuiltInParameter.SHEET_NAME)
            }, warnings);
#else
            if (!source.CanBeDuplicated(SheetDuplicateOption.DuplicateSheetWithDetailing))
                throw new InvalidOperationException("Revit не разрешает дублировать этот лист с оформлением.");
            sheet = (ViewSheet)doc.GetElement(source.Duplicate(SheetDuplicateOption.DuplicateSheetWithDetailing));

            // Обновляем состав дубликата до чтения созданных Revit окон.
            doc.Regenerate();

            // Штатное дублирование уже переносит аннотации в исходных координатах листа.
            // При отключённой опции удаляем их только из нового листа, не затрагивая штамп.
            if (!model.CopySheetAnnotations)
            {
                reportStage("исключение аннотаций листа");
                RemoveSheetAnnotations(doc, sheet);
            }

            // Общие легенды сохраняем вместе с окнами: удалять и создавать их повторно не требуется.
            var sharedLegendIds = new HashSet<ElementId>(model.Contents
                .Where(c => c.Action == CopyActions.Shared && c.View.ViewType == ViewType.Legend)
                .Select(c => c.View.Id));
            var placements = new List<ElementId>();
            foreach (Viewport viewport in sheet.GetAllViewports().Select(doc.GetElement).OfType<Viewport>())
            {
                if (sharedLegendIds.Contains(viewport.ViewId) && !retainedViewports.ContainsKey(viewport.ViewId))
                    retainedViewports.Add(viewport.ViewId, viewport);
                else
                    placements.Add(viewport.Id);
            }
            placements.AddRange(new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance)).Cast<ScheduleSheetInstance>()
                .Where(i => i.OwnerViewId == sheet.Id && !i.IsTitleblockRevisionSchedule).Select(i => i.Id));
            if (placements.Count > 0)
            {
                doc.Delete(placements);
                doc.Regenerate();
            }
#endif
            sheet.SheetNumber = row.Number;
            sheet.Name = row.Name;

            foreach (ContentRow content in model.Contents)
            {
                if (content.Action == CopyActions.Skip) continue;
                reportStage("вид «" + content.View.Name + "»");
                View target;
                if (content.Action == CopyActions.NewPlan)
                {
                    string name = content.View.Id == model.MainPlan.Id ? row.ViewName
                        : RevitHelpers.UniqueViewName(doc, content.View.Name + " — " + row.Level.Name);
                    target = PlanService.Create(doc, (ViewPlan)content.View, row, name, model.CopyPlanDetails, warnings);
                }
                else if (content.Action == CopyActions.Duplicate)
                {
                    target = (View)doc.GetElement(content.View.Duplicate(ViewDuplicateOption.WithDetailing));
                    target.Name = RevitHelpers.UniqueViewName(doc, content.View.Name + " — " + row.Level.Name);
                    if (content.View is ViewPlan || content.View is ViewSection)
                        warnings.Add("Дополнительный вид «" + content.View.Name + "» скопирован в исходной области модели; на другой этаж не переносился.");
                }
                else if (content.Action == CopyActions.Replace)
                    target = (View)doc.GetElement(content.Replacement.Id);
                else target = content.View;

                retainedViewports.TryGetValue(target.Id, out Viewport retainedViewport);
                PlaceView(doc, sheet, content.Viewport, target, warnings, retainedViewport);
            }

            foreach (ScheduleRow schedule in model.Schedules.Where(s => s.Include))
            {
                reportStage("спецификация «" + schedule.Schedule.Name + "»");
                ViewSchedule target = ScheduleService.Create(doc, schedule, row.Level, row.Index, warnings);
#if !Debug2020 && !Revit2020
                if (schedule.Copy)
                    RestoreScheduleSegments(doc, schedule.Schedule, target);

                if (target.GetSegmentCount() > 1 && schedule.Instances.All(instance => instance.SegmentIndex >= 0)
                    && schedule.Instances.Select(instance => instance.SegmentIndex).Distinct().Count() < target.GetSegmentCount())
                    warnings.Add("На листе-эталоне размещены не все части спецификации «" + target.Name
                        + "». На новый лист перенесены только части с этого листа, в исходных позициях.");
#endif
                foreach (ScheduleSheetInstance original in schedule.Instances)
                {
#if Debug2020 || Revit2020
                    var placed = ScheduleSheetInstance.Create(doc, sheet.Id, target.Id, original.Point);
#else
                    // -1 размещает всю таблицу; 0 и выше — конкретную часть эталона.
                    var placed = ScheduleSheetInstance.Create(doc, sheet.Id, target.Id, original.Point, original.SegmentIndex);
#endif
                    placed.Rotation = original.Rotation;
                }
            }
            doc.Regenerate();
            reportStage("проверка размещения");
            CheckLayout(doc, sheet, warnings);
            return sheet;
        }

#if !Debug2020 && !Revit2020
        private static void RestoreScheduleSegments(Document doc, ViewSchedule source, ViewSchedule target)
        {
            // После замены фильтров меняется число строк, но границы частей остаются как у эталона.
            // Меняем только независимую копию; исходная спецификация не затрагивается.
            int count = source.GetSegmentCount();
            if (target.GetSegmentCount() != count)
            {
                while (target.GetSegmentCount() > 1)
                    target.MergeSegments(target.GetSegmentCount() - 1, target.GetSegmentCount() - 2);

                if (count > 1)
                    target.Split(count);
            }

            // Высота последней части не задаётся: Revit размещает в ней оставшиеся строки.
            for (int index = 0; index < count - 1; index++)
                target.SetSegmentHeight(index, source.GetSegmentHeight(index));
            doc.Regenerate();
        }
#endif

        private static bool IsSheetAnnotation(Element element)
        {
            if (element is Viewport || element is ScheduleSheetInstance || element is View
                || element.Category?.Id == new ElementId(BuiltInCategory.OST_TitleBlocks))
                return false;

            return element.Category?.CategoryType == CategoryType.Annotation
                || element is TextNote || element is DetailCurve || element is FilledRegion
                || element is ImageInstance
                || (element is Group && element.Category?.Id == new ElementId(BuiltInCategory.OST_IOSDetailGroups));
        }

        private static void RemoveSheetAnnotations(Document doc, ViewSheet sheet)
        {
            var annotations = new FilteredElementCollector(doc).OwnedByView(sheet.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.GroupId == ElementId.InvalidElementId && IsSheetAnnotation(e))
                .Select(e => e.Id).ToList();
            if (annotations.Count == 0)
                return;

            doc.Delete(annotations);
            doc.Regenerate();
        }

        private static void PlaceView(Document doc, ViewSheet sheet, Viewport source, View view, List<string> warnings,
            Viewport placed = null)
        {
            if (placed == null)
            {
                if (!Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                    throw new InvalidOperationException("Вид нельзя разместить на листе (возможно, уже размещён): " + view.Name);

                placed = Viewport.Create(doc, sheet.Id, view.Id, source.GetBoxCenter());
                if (placed == null)
                    throw new InvalidOperationException("Revit не создал окно вида «" + view.Name + "» [" + view.Id
                        + ", " + view.ViewType + "] на листе «" + sheet.SheetNumber
                        + "», хотя предварительная проверка разрешила размещение. Создание листа будет отменено.");
            }

            placed.ChangeTypeId(source.GetTypeId());
            placed.Rotation = source.Rotation;
            doc.Regenerate();

#if !Debug2020 && !Revit2020
            placed.LabelOffset = source.LabelOffset;
            placed.LabelLineLength = source.LabelLineLength;
            if (view is ViewPlan && doc.GetElement(source.ViewId) is ViewPlan)
            {
                // Совмещаем одну модельную точку в координатах листа, а не центр меняющегося габарита.
                View fromView = (View)doc.GetElement(source.ViewId);
                var fromTransforms = fromView.GetModelToProjectionTransforms();
                var toTransforms = view.GetModelToProjectionTransforms();
                if (fromTransforms.Count == 1 && toTransforms.Count == 1)
                {
                    XYZ anchor = FindAnchor((ViewPlan)fromView);
                    XYZ targetAnchor = new XYZ(anchor.X, anchor.Y, ((ViewPlan)view).GenLevel.ProjectElevation);
                    XYZ from = source.GetProjectionToSheetTransform().OfPoint(fromTransforms[0].GetModelToProjectionTransform().OfPoint(anchor));
                    XYZ to = placed.GetProjectionToSheetTransform().OfPoint(toTransforms[0].GetModelToProjectionTransform().OfPoint(targetAnchor));
                    placed.SetBoxCenter(placed.GetBoxCenter() + new XYZ(from.X - to.X, from.Y - to.Y, 0));
                }
                else warnings.Add("У вида несколько областей проекции; сохранён центр окна: " + view.Name);
            }
            else placed.SetBoxCenter(source.GetBoxCenter());
#else
            placed.SetBoxCenter(source.GetBoxCenter());
#endif
            doc.Regenerate();
        }

        private static void CheckLayout(Document doc, ViewSheet sheet, List<string> warnings)
        {
            var bounds = new List<Tuple<string, double, double, double, double>>();
            foreach (Viewport viewport in sheet.GetAllViewports().Select(doc.GetElement).OfType<Viewport>())
            {
                Outline box = viewport.GetBoxOutline();
                Outline label = viewport.GetLabelOutline();
                if (label.MaximumPoint.DistanceTo(label.MinimumPoint) < 1e-8) label = box;
                bounds.Add(Tuple.Create(doc.GetElement(viewport.ViewId).Name,
                    Math.Min(box.MinimumPoint.X, label.MinimumPoint.X), Math.Min(box.MinimumPoint.Y, label.MinimumPoint.Y),
                    Math.Max(box.MaximumPoint.X, label.MaximumPoint.X), Math.Max(box.MaximumPoint.Y, label.MaximumPoint.Y)));
            }
            foreach (ScheduleSheetInstance instance in new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>().Where(i => i.OwnerViewId == sheet.Id && !i.IsTitleblockRevisionSchedule))
            {
                BoundingBoxXYZ box = instance.get_BoundingBox(sheet);
                if (box != null) bounds.Add(Tuple.Create(doc.GetElement(instance.ScheduleId).Name, box.Min.X, box.Min.Y, box.Max.X, box.Max.Y));
            }
            BoundingBoxUV paper = sheet.Outline;
            double tolerance = 1.0 / 304.8;
            for (int i = 0; i < bounds.Count; i++)
            {
                var a = bounds[i];
                if (a.Item2 < paper.Min.U - tolerance || a.Item3 < paper.Min.V - tolerance
                    || a.Item4 > paper.Max.U + tolerance || a.Item5 > paper.Max.V + tolerance)
                    warnings.Add("Габарит выходит за лист: " + a.Item1);
                for (int j = i + 1; j < bounds.Count; j++)
                {
                    var b = bounds[j];
                    if (Math.Min(a.Item4, b.Item4) - Math.Max(a.Item2, b.Item2) > tolerance
                        && Math.Min(a.Item5, b.Item5) - Math.Max(a.Item3, b.Item3) > tolerance)
                        warnings.Add("Проверьте пересечение габаритов: " + a.Item1 + " / " + b.Item1);
                }
            }
        }

        private static XYZ FindAnchor(ViewPlan view)
        {
            using (var crop = view.GetCropRegionShapeManager())
            {
                var loop = crop.GetCropShape().FirstOrDefault();
                if (!view.CropBoxActive || loop == null) return new XYZ(0, 0, view.GenLevel.ProjectElevation);
                var points = loop.SelectMany(c => c.Tessellate()).ToList();
                double minX = points.Min(p => p.X), maxX = points.Max(p => p.X);
                double minY = points.Min(p => p.Y), maxY = points.Max(p => p.Y);
                for (int ix = 1; ix < 20; ix++)
                for (int iy = 1; iy < 20; iy++)
                {
                    double x = minX + (maxX - minX) * ix / 20.0;
                    double y = minY + (maxY - minY) * iy / 20.0;
                    bool inside = false;
                    for (int i = 0, j = points.Count - 1; i < points.Count; j = i++)
                    {
                        XYZ a = points[i], b = points[j];
                        if ((a.Y > y) != (b.Y > y) && x < (b.X - a.X) * (y - a.Y) / (b.Y - a.Y) + a.X)
                            inside = !inside;
                    }
                    if (inside) return new XYZ(x, y, view.GenLevel.ProjectElevation);
                }
                throw new InvalidOperationException("Не удалось найти опорную точку внутри обрезки: " + view.Name);
            }
        }

        private sealed class GenerationFailures : IFailuresPreprocessor
        {
            private readonly List<string> _messages;
            internal GenerationFailures(List<string> messages) { _messages = messages; }

            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                bool errors = false;
                foreach (FailureMessageAccessor message in accessor.GetFailureMessages())
                {
                    _messages.Add(message.GetDescriptionText());
                    if (message.GetSeverity() == FailureSeverity.Warning) accessor.DeleteWarning(message);
                    else errors = true;
                }
                return errors ? FailureProcessingResult.ProceedWithRollBack : FailureProcessingResult.Continue;
            }
        }
    }
}
