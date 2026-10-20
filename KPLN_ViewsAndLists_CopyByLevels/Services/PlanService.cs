using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using KPLN_ViewsAndLists_Ribbon.Common.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ViewsAndLists_CopyByLevels.Services
{
    internal static class PlanService
    {
        private static readonly PlanViewPlane[] Planes =
        {
            PlanViewPlane.TopClipPlane, PlanViewPlane.CutPlane,
            PlanViewPlane.BottomClipPlane, PlanViewPlane.ViewDepthPlane
        };

        internal static ViewPlan Create(Document doc, ViewPlan source, PreviewRow row, string name, bool details, List<string> warnings)
        {
            if (source.GenLevel == null || source.ViewType == ViewType.AreaPlan)
                throw new InvalidOperationException("Для поэтажного создания нужен план этажа, потолка или несущих конструкций.");

            ViewPlan target = ViewPlan.Create(doc, source.GetTypeId(), row.Level.Level.Id);
            target.Name = name;
            target.ViewTemplateId = source.ViewTemplateId;
            target.ApplyViewTemplateParameters(source);
            doc.Regenerate();

            if (target.Scale != source.Scale)
                throw new InvalidOperationException("Шаблон не позволил сохранить масштаб исходного плана.");
            if (!target.RightDirection.IsAlmostEqualTo(source.RightDirection) || !target.UpDirection.IsAlmostEqualTo(source.UpDirection))
                throw new InvalidOperationException("У эталона повернута ориентация плана. В этой бете нужен эталон с обычной ориентацией или областью видимости.");

            CopyRange(doc, source, target, warnings);
            CopyCrop(doc, source, target);

            // Заголовок должен соответствовать новому этажу, а не свободному тексту исходного вида.
            Parameter title = target.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION);
            if (title != null && !title.IsReadOnly) title.Set(row.Name);
            if (details) CopyIndependentDetails(doc, source, target, warnings);
            int bound = new FilteredElementCollector(doc).OwnedByView(source.Id)
                .Count(e => e is IndependentTag || e is SpatialElementTag || e is Dimension);
            if (bound > 0) warnings.Add(source.Name + ": марки и размеры с привязками не переносились (" + bound + "). Нужна разметка нового этажа.");

            doc.Regenerate();
            if (target.GenLevel.Id != row.Level.Level.Id || target.ViewTemplateId != source.ViewTemplateId)
                throw new InvalidOperationException("Не подтверждены уровень или шаблон нового плана.");
            return target;
        }

        private static void CopyRange(Document doc, ViewPlan source, ViewPlan target, List<string> warnings)
        {
            using (PlanViewRange expected = source.GetViewRange())
            {
                foreach (PlanViewPlane plane in Planes)
                {
                    ElementId referenceId = expected.GetLevelId(plane);
                    var reference = doc.GetElement(referenceId) as Level;
                    if (reference != null)
                    {
                        // Явная ссылка на уровень становится эквивалентным смещением от нового этажа.
                        double offset = expected.GetOffset(plane) + reference.ProjectElevation - source.GenLevel.ProjectElevation;
                        expected.SetLevelId(plane, target.GenLevel.Id);
                        expected.SetOffset(plane, offset);
                        if (reference.Id != source.GenLevel.Id)
                            warnings.Add("Диапазон «" + plane + "»: явный уровень «" + reference.Name + "» пересчитан относительно целевого этажа.");
                    }
                    // Специальные ссылки «связанный/выше/ниже/неограниченно» остаются относительными.
                }

                if (!RevitHelpers.TemplateControls(doc, target, BuiltInParameter.PLAN_VIEW_RANGE))
                {
                    if (target.CheckPlanViewRangeValidity(expected).Count > 0)
                        throw new InvalidOperationException("Диапазон эталона недопустим для нового плана.");
                    target.SetViewRange(expected);
                    doc.Regenerate();
                }

                using (PlanViewRange actual = target.GetViewRange())
                {
                    if (Planes.Any(p => actual.GetLevelId(p) != expected.GetLevelId(p)
                        || Math.Abs(actual.GetOffset(p) - expected.GetOffset(p)) > 1e-8))
                        throw new InvalidOperationException("Диапазон не перенесён. Проверьте управление диапазоном в шаблоне вида; общий шаблон не изменён.");
                }
            }
        }

        private static void CopyCrop(Document doc, ViewPlan source, ViewPlan target)
        {
            Parameter sourceScope = source.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
            ElementId scopeId = sourceScope?.AsElementId() ?? ElementId.InvalidElementId;
            if (scopeId != ElementId.InvalidElementId)
            {
                Parameter scope = target.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP);
                if (scope != null && !scope.IsReadOnly) scope.Set(scopeId);
                if (scope == null || scope.AsElementId() != scopeId)
                    throw new InvalidOperationException("Не удалось перенести область видимости.");
            }
            else
            {
                using (ViewCropRegionShapeManager from = source.GetCropRegionShapeManager())
                {
                    if (from.NumberOfSplitRegions > 1)
                        throw new InvalidOperationException("Разделённая обрезка плана пока не поддерживается. Выберите эталон с цельной обрезкой.");
                }
            }

            // Общий алгоритм с командой «Копировать подрезку»; область видимости сохраняет свой контур.
            doc.Regenerate();
            ViewCropCopyService.Copy(source, target, scopeId == ElementId.InvalidElementId);
        }

        private static void CopyIndependentDetails(Document doc, View source, View target, List<string> warnings)
        {
            var elements = new FilteredElementCollector(doc).OwnedByView(source.Id).WhereElementIsNotElementType().ToList();
            foreach (Element element in elements)
            {
                bool independent = element is TextNote || element is DetailCurve || element is FilledRegion
                    || element is ImageInstance || (element is FamilyInstance && element.Category?.Id == new ElementId(BuiltInCategory.OST_GenericAnnotation));
                if (!independent) continue;
                using (var sub = new SubTransaction(doc))
                {
                    sub.Start();
                    try
                    {
                        ElementTransformUtils.CopyElements(source, new[] { element.Id }, target, Transform.Identity, new CopyPasteOptions());
                        sub.Commit();
                    }
                    catch (Exception ex)
                    {
                        sub.RollBack();
                        warnings.Add("Оформление вида, ID " + element.Id + ": " + ex.Message);
                    }
                }
            }
        }
    }
}
