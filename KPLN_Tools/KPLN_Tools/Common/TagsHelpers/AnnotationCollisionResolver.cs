using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.Common.TagsHelpers
{
    /// <summary>
    /// Универсальный класс для проверки коллизий аннотаций на виде.
    /// Можно переиспользовать в других плагинах (марки, размеры, тексты).
    ///
    internal sealed class AnnotationCollisionResolver
    {
        private readonly Document _doc;
        private readonly View _view;

        /// <summary>Допуск, на который BBox-ы могут «коснуться» без коллизии (в футах).</summary>
#if Debug2020 || Revit2020
        public double Tolerance { get; set; } = UnitUtils.ConvertToInternalUnits(2.0, DisplayUnitType.DUT_MILLIMETERS);
#else
        public double Tolerance { get; set; } = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters);
#endif

        /// <summary>
        /// Категории, считающиеся «аннотациями» для проверки коллизий.
        /// При необходимости можно расширить или сузить извне.
        /// </summary>
        public HashSet<BuiltInCategory> AnnotationCategories { get; } = new HashSet<BuiltInCategory>
        {
            // Tags
            BuiltInCategory.OST_PipeTags,
            BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_MEPSpaceTags,
            BuiltInCategory.OST_RoomTags,
            BuiltInCategory.OST_DoorTags,
            BuiltInCategory.OST_WindowTags,
            BuiltInCategory.OST_WallTags,
            BuiltInCategory.OST_FloorTags,
            BuiltInCategory.OST_GenericModelTags,
            BuiltInCategory.OST_PipeFittingTags,
            BuiltInCategory.OST_PipeAccessoryTags,
            BuiltInCategory.OST_DuctFittingTags,
            BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_MultiCategoryTags,
            BuiltInCategory.OST_PipeInsulationsTags,
            BuiltInCategory.OST_DuctInsulationsTags,
            // Размеры
            BuiltInCategory.OST_Dimensions,
            // Текст
            BuiltInCategory.OST_TextNotes,
            // Условные обозначения
            BuiltInCategory.OST_GenericAnnotation,
            // Спецификации/легенды как элементы вида не учитываются (отдельные виды)
        };

        public AnnotationCollisionResolver(Document doc, View view)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _view = view ?? throw new ArgumentNullException(nameof(view));
        }

        /// <summary>
        /// Проверяет, пересекается ли указанный элемент с любой другой
        /// аннотацией на виде.
        /// </summary>
        public bool HasCollision(Element element)
        {
            return GetCollidingElements(element).Any();
        }

        /// <summary>
        /// Возвращает список элементов, пересекающихся с указанным.
        /// </summary>
        public IEnumerable<Element> GetCollidingElements(Element element)
        {
            BoundingBoxXYZ targetBox = element.get_BoundingBox(_view);
            if (targetBox == null) 
                yield break;

            var others = CollectAnnotations()
                .Where(e => e.Id != element.Id);

            foreach (var other in others)
            {
                var otherBox = other.get_BoundingBox(_view);
                if (otherBox == null) continue;

                if (BoxesIntersectXY(targetBox, otherBox, Tolerance))
                    yield return other;
            }
        }

        /// <summary>
        /// Сбор всех аннотаций на текущем виде по списку категорий.
        /// </summary>
        private IEnumerable<Element> CollectAnnotations()
        {
            var catFilters = AnnotationCategories
                .Select(c => (ElementFilter)new ElementCategoryFilter(c))
                .ToList();

            ElementFilter filter = catFilters.Count == 1
                ? catFilters[0]
                : new LogicalOrFilter(catFilters);

            return new FilteredElementCollector(_doc, _view.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();
        }

        /// <summary>
        /// Проверка пересечения BBox-ов в плоскости XY с допуском.
        /// </summary>
        private static bool BoxesIntersectXY(BoundingBoxXYZ a, BoundingBoxXYZ b, double tolerance)
        {
            // Если BBox в локальной системе, приводим к мировой через Transform
            (XYZ aMin, XYZ aMax) = ToWorldMinMax(a);
            (XYZ bMin, XYZ bMax) = ToWorldMinMax(b);

            // Сжимаем «прикасание» на tolerance — две марки, чьи бордюры впритык,
            // считаем НЕ пересекающимися.
            bool sepX = (aMax.X - tolerance) <= bMin.X || (bMax.X - tolerance) <= aMin.X;
            bool sepY = (aMax.Y - tolerance) <= bMin.Y || (bMax.Y - tolerance) <= aMin.Y;

            return !(sepX || sepY);
        }

        private static (XYZ min, XYZ max) ToWorldMinMax(BoundingBoxXYZ box)
        {
            Transform t = box.Transform ?? Transform.Identity;

            // 8 углов локального BBox -> в мировые координаты -> ищем min/max.
            XYZ[] corners = {
                new XYZ(box.Min.X, box.Min.Y, box.Min.Z),
                new XYZ(box.Min.X, box.Min.Y, box.Max.Z),
                new XYZ(box.Min.X, box.Max.Y, box.Min.Z),
                new XYZ(box.Min.X, box.Max.Y, box.Max.Z),
                new XYZ(box.Max.X, box.Min.Y, box.Min.Z),
                new XYZ(box.Max.X, box.Min.Y, box.Max.Z),
                new XYZ(box.Max.X, box.Max.Y, box.Min.Z),
                new XYZ(box.Max.X, box.Max.Y, box.Max.Z),
            };

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            foreach (var c in corners)
            {
                XYZ w = t.OfPoint(c);
                if (w.X < minX) minX = w.X; if (w.X > maxX) maxX = w.X;
                if (w.Y < minY) minY = w.Y; if (w.Y > maxY) maxY = w.Y;
                if (w.Z < minZ) minZ = w.Z; if (w.Z > maxZ) maxZ = w.Z;
            }

            return (new XYZ(minX, minY, minZ), new XYZ(maxX, maxY, maxZ));
        }
    }
}
