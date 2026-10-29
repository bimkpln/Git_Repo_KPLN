using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Publication
{
    internal sealed class SheetBoundaryIssue
    {
        internal ViewSheet Sheet { get; }
        internal IList<ElementId> ElementIds { get; }
        internal string Details { get; }

        internal SheetBoundaryIssue(ViewSheet sheet, IList<Element> elements)
        {
            Sheet = sheet;
            ElementIds = elements.Select(element => element.Id).ToList();
            Details = "ID для выделения: " + string.Join(", ", ElementIds.Select(id => id.ToString()).Distinct()) + "\n\n"
                + string.Join("\n", elements.Take(20).Select(element =>
                $"{element.Category?.Name ?? "Элемент"}: {element.Name} (Id {element.Id})"));
            if (elements.Count > 20)
                Details += $"\nЕщё элементов: {elements.Count - 20}.";
        }
    }

    internal static class SheetBoundaryCheck
    {
        internal static SheetBoundaryIssue Find(MainEntity entity)
        {
            // Несколько штампов — существующий режим печати листа по частям.
            if (!(entity.MainView is ViewSheet sheet) || entity.TitleBlocks == null
                || entity.TitleBlocks.Count != 1)
                return null;

            FamilyInstance titleBlock = entity.TitleBlocks[0];
            BoundingBoxXYZ frameBox = titleBlock.get_BoundingBox(sheet)
                ?? throw new InvalidOperationException("Не удалось проверить границы основной надписи: " + entity);
            SheetPrintBounds frame = GetBounds(frameBox);
            var outside = new List<Element>();

            // Коллектор вида учитывает видимость и фильтры. Дополнительно исключаем
            // постоянное/временное скрытие и категории, скрываемые профилем печати.
            using (var collector = new FilteredElementCollector(sheet.Document, sheet.Id))
            {
                foreach (Element element in collector.WhereElementIsNotElementType())
                {
                    if (element.Id.Equals(titleBlock.Id) || !IsPrinted(sheet, element))
                        continue;

                    Viewport viewport = element as Viewport;
                    if (viewport != null)
                    {
                        Outline box = viewport.GetBoxOutline();
                        Outline label = viewport.GetLabelOutline();
                        bool labelOutside = label != null && !label.MinimumPoint.IsAlmostEqualTo(label.MaximumPoint)
                            && GetBounds(label).ExtendsOutside(frame);
                        if (GetBounds(box).ExtendsOutside(frame) || labelOutside)
                            outside.Add(element);
                        continue;
                    }

                    BoundingBoxXYZ bounds = element.get_BoundingBox(sheet);
                    if (bounds != null && GetBounds(bounds).ExtendsOutside(frame))
                        outside.Add(element);
                }
            }

            return outside.Count == 0 ? null : new SheetBoundaryIssue(sheet, outside);
        }

        private static bool IsPrinted(ViewSheet sheet, Element element)
        {
            // Элементы размещённых видов проверяются через габарит их видового экрана.
            if (!element.OwnerViewId.Equals(sheet.Id) && !(element is Viewport))
                return false;
            if (element.IsHidden(sheet))
                return false;

            Category category = element.Category;
            if (category != null)
            {
                if (category.Id.Equals(new ElementId(BuiltInCategory.OST_GuideGrid))
                    || category.Id.Equals(new ElementId(BuiltInCategory.OST_VolumeOfInterest))
                    || category.Id.Equals(new ElementId(BuiltInCategory.OST_CLines)))
                    return false;
                for (Category current = category; current != null; current = current.Parent)
                {
                    if (sheet.CanCategoryBeHidden(current.Id) && sheet.GetCategoryHidden(current.Id))
                        return false;
                }
            }

            return !sheet.IsInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
                || sheet.IsElementVisibleInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate, element.Id);
        }

        private static SheetPrintBounds GetBounds(Outline box)
        {
            return new SheetPrintBounds(box.MinimumPoint.X, box.MinimumPoint.Y, box.MaximumPoint.X, box.MaximumPoint.Y);
        }

        private static SheetPrintBounds GetBounds(BoundingBoxXYZ box)
        {
            var points = new List<XYZ>();
            foreach (double x in new[] { box.Min.X, box.Max.X })
            foreach (double y in new[] { box.Min.Y, box.Max.Y })
            foreach (double z in new[] { box.Min.Z, box.Max.Z })
                points.Add(box.Transform.OfPoint(new XYZ(x, y, z)));

            return new SheetPrintBounds(points.Min(p => p.X), points.Min(p => p.Y),
                points.Max(p => p.X), points.Max(p => p.Y));
        }
    }
}
