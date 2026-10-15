using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class PlanRoutingService
    {
        private static readonly PlanViewPlane[] Planes = {
            PlanViewPlane.TopClipPlane, PlanViewPlane.CutPlane,
            PlanViewPlane.BottomClipPlane, PlanViewPlane.ViewDepthPlane
        };

        internal static object ReadRange(Document doc, IDictionary<string, object> input)
        {
            var plan = DocumentationReadService.Require<ViewPlan>(doc, input, "view_unique_id");
            using (var range = plan.GetViewRange())
            {
                return new {
                    view = RevitService.ElementInfo(plan),
                    level = plan.GenLevel == null ? null : RevitService.ElementInfo(plan.GenLevel),
                    template_id = RevitService.IdValue(plan.ViewTemplateId).ToString(),
                    planes = Planes.Select(p => new {
                        plane = p.ToString(), level_id = RevitService.IdValue(range.GetLevelId(p)).ToString(),
                        level_name = doc.GetElement(range.GetLevelId(p))?.Name,
                        offset_mm = range.GetOffset(p) * 304.8
                    }).ToArray()
                };
            }
        }

        internal static Element CopyRange(Document doc, IDictionary<string, object> input)
        {
            var target = DocumentationReadService.Require<ViewPlan>(doc, input, "view_unique_id");
            var source = DocumentationReadService.Require<ViewPlan>(doc, input, "source_view_unique_id");
            if (target.IsTemplate || source.IsTemplate || target.Id == source.Id)
                throw new BridgeException("invalid_view", "Нужны два разных плана, не шаблоны.");

            // Для этого действия источник обязан относиться к тому же этажу.
            if (target.GenLevel == null || source.GenLevel == null || target.GenLevel.Id != source.GenLevel.Id)
                throw new BridgeException("different_level", "Планы должны иметь один связанный уровень.");

            var template = doc.GetElement(target.ViewTemplateId) as View;
            var rangeParameter = new ElementId(BuiltInParameter.PLAN_VIEW_RANGE);
            if (template != null && template.GetTemplateParameterIds().Contains(rangeParameter)
                && !template.GetNonControlledTemplateParameterIds().Contains(rangeParameter))
                throw new BridgeException("template_controls_range", "Секущий диапазон управляется шаблоном; общий шаблон не изменён.");

            using (var range = source.GetViewRange())
            {
                target.SetViewRange(range);
                doc.Regenerate();
                using (var actual = target.GetViewRange())
                {
                    if (Planes.Any(p => actual.GetLevelId(p) != range.GetLevelId(p)
                        || Math.Abs(actual.GetOffset(p) - range.GetOffset(p)) > 1e-8))
                        throw new BridgeException("range_not_applied", "Revit не сохранил запрошенный секущий диапазон.");
                }
            }
            return target;
        }

        internal static Element SetRange(Document doc, IDictionary<string, object> input)
        {
            var target = DocumentationReadService.Require<ViewPlan>(doc, input, "view_unique_id");
            if (target.IsTemplate || target.GenLevel == null)
                throw new BridgeException("invalid_view", "Нужен план со связанным уровнем, не шаблон.");

            var template = doc.GetElement(target.ViewTemplateId) as View;
            var rangeParameter = new ElementId(BuiltInParameter.PLAN_VIEW_RANGE);
            if (template != null && template.GetTemplateParameterIds().Contains(rangeParameter)
                && !template.GetNonControlledTemplateParameterIds().Contains(rangeParameter))
                throw new BridgeException("template_controls_range", "Секущий диапазон управляется шаблоном; общий шаблон не изменён.");


            // Все четыре смещения задаются явно от связанного уровня целевого плана.
            var keys = new[] { "top_offset_mm", "cut_offset_mm", "bottom_offset_mm", "depth_offset_mm" };
            var offsets = keys.Select(key => Json.Double(input.Get(key)) / 304.8).ToArray();
            using (var range = target.GetViewRange())
            {
                for (int i = 0; i < Planes.Length; i++)
                {
                    range.SetLevelId(Planes[i], target.GenLevel.Id);
                    range.SetOffset(Planes[i], offsets[i]);
                }

                var errors = target.CheckPlanViewRangeValidity(range);
                if (errors.Count > 0)
                    throw new BridgeException("invalid_plan_range", "Недопустимый секущий диапазон: " + string.Join("; ", errors));

                target.SetViewRange(range);
                doc.Regenerate();

                // Подтверждаем фактическое применение до фиксации общей транзакции.
                using (var actual = target.GetViewRange())
                {
                    if (Planes.Any(p => actual.GetLevelId(p) != range.GetLevelId(p)
                        || Math.Abs(actual.GetOffset(p) - range.GetOffset(p)) > 1e-8))
                        throw new BridgeException("range_not_applied", "Revit не сохранил запрошенный секущий диапазон.");
                }
            }
            return target;
        }

        internal static Element CopyCrop(Document doc, IDictionary<string, object> input)
        {
            var target = DocumentationReadService.Require<ViewPlan>(doc, input, "view_unique_id");
            var source = DocumentationReadService.Require<ViewPlan>(doc, input, "source_view_unique_id");
            if (target.Id == source.Id || target.IsTemplate || source.IsTemplate
                || target.GenLevel == null || source.GenLevel == null || !source.CropBoxActive)
                throw new BridgeException("invalid_view", "Нужны разные планы с уровнями и включённой обрезкой образца.");
            if (!target.ViewDirection.IsAlmostEqualTo(source.ViewDirection)
                || !target.RightDirection.IsAlmostEqualTo(source.RightDirection)
                || !target.UpDirection.IsAlmostEqualTo(source.UpDirection))
                throw new BridgeException("different_orientation", "Ориентации планов должны совпадать.");
            if (target.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)?.AsElementId() != ElementId.InvalidElementId)
                throw new BridgeException("scope_box_controls_crop", "Обрезка целевого плана связана с областью видимости.");


            using (var from = source.GetCropRegionShapeManager())
            using (var to = target.GetCropRegionShapeManager())
            {
                var loops = from.GetCropShape();
                if (from.NumberOfSplitRegions > 1 || to.NumberOfSplitRegions > 1 || loops.Count != 1 || !to.CanHaveShape)
                    throw new BridgeException("unsupported_crop", "Поддерживается один цельный контур обрезки без разделения.");

                var shift = Transform.CreateTranslation(XYZ.BasisZ * (target.GenLevel.ProjectElevation - source.GenLevel.ProjectElevation));
                var contour = CurveLoop.CreateViaTransform(loops[0], shift);
                if (!to.IsCropRegionShapeValid(contour))
                    throw new BridgeException("invalid_crop", "Контур образца недопустим для целевого плана.");

                target.CropBoxActive = true;
                to.SetCropShape(contour);
                target.CropBoxVisible = source.CropBoxVisible;
                doc.Regenerate();

                var expected = contour.Select(c => c.GetEndPoint(0)).ToArray();
                var actualLoops = to.GetCropShape();
                var actual = actualLoops.Count == 1 ? actualLoops[0].Select(c => c.GetEndPoint(0)).ToArray() : new XYZ[0];
                if (actual.Length != expected.Length || expected.Any(p => !actual.Any(q =>
                    AnnotationPlane.Distance(p, q, target.ViewDirection) <= 0.1 / 304.8)))
                    throw new BridgeException("crop_not_applied", "Revit не подтвердил контур обрезки; пакет отменён.");
            }
            return target;
        }

        internal static object Connections(Document doc, IDictionary<string, object> input)
        {
            var ids = input.List("unique_ids", 100).Cast<object>().Select(x => x as string).ToArray();
            if (ids.Any(string.IsNullOrEmpty))
                throw new BridgeException("invalid_input", "Нужны UniqueId элементов.");

            var elements = ids.Select(id => doc.GetElement(id)
                ?? throw new BridgeException("not_found", "Элемент не найден: " + id)).ToArray();
            return elements.Select(element => {
                var curve = element as MEPCurve;
                var family = element as FamilyInstance;
                var manager = curve?.ConnectorManager ?? family?.MEPModel?.ConnectorManager;
                if (manager == null)
                    throw new BridgeException("no_connectors", "У элемента нет MEP-соединителей: " + element.UniqueId);

                return new {
                    element = RevitService.ElementInfo(element),
                    part_type = (family?.MEPModel as MechanicalFitting)?.PartType.ToString(),
                    connectors = manager.Connectors.Cast<Connector>()
                        .Where(c => c.ConnectorType == ConnectorType.End || c.ConnectorType == ConnectorType.Curve)
                        .Select(c => new {
                            id = c.Id, origin_mm = new[] { c.Origin.X * 304.8, c.Origin.Y * 304.8, c.Origin.Z * 304.8 },
                            connector_type = c.ConnectorType.ToString(), domain = c.Domain.ToString(),
                            connected_to = c.AllRefs.Cast<Connector>()
                                .Where(r => r.Owner.Id != element.Id
                                    && (r.ConnectorType == ConnectorType.End || r.ConnectorType == ConnectorType.Curve)
                                    && c.IsConnectedTo(r))
                                .Select(r => new { element = RevitService.ElementInfo(r.Owner), connector_id = r.Id }).ToArray()
                        }).ToArray()
                };
            }).ToArray();
        }
    }
}
