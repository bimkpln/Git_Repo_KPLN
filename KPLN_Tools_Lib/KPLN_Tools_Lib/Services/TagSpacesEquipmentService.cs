using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_Lib.Services
{
    internal sealed class TagSpacesEquipmentOptions
    {
        internal TagSpacesEquipmentOptions(SpaceTagType radiatorType, SpaceTagType grilleType)
        {
            RadiatorType = radiatorType;
            GrilleType = grilleType;
        }

        internal SpaceTagType RadiatorType { get; }
        internal SpaceTagType GrilleType { get; }
        internal bool Enabled => RadiatorType != null || GrilleType != null;
    }

    internal sealed class TagSpacesEquipmentAnalysis
    {
        internal HashSet<ElementId> RadiatorSpaces { get; } = new HashSet<ElementId>();
        internal HashSet<ElementId> GrilleSpaces { get; } = new HashSet<ElementId>();
        internal List<string> Messages { get; } = new List<string>();
        internal string LayoutError { get; set; }

        internal FamilySymbol GetTagType(ElementId spaceId, FamilySymbol defaultType, TagSpacesEquipmentOptions options)
        {
            if (options.RadiatorType != null && RadiatorSpaces.Contains(spaceId))
                return options.RadiatorType;
            if (options.GrilleType != null && GrilleSpaces.Contains(spaceId))
                return options.GrilleType;
            return defaultType;
        }
    }

    internal sealed class TagSpacesEquipmentService
    {
#if Debug2020 || Revit2020
        private static readonly double MaxGrilleProjection = UnitUtils.ConvertToInternalUnits(1500, DisplayUnitType.DUT_MILLIMETERS);
        private static readonly double GrilleHalfWidth = UnitUtils.ConvertToInternalUnits(25, DisplayUnitType.DUT_MILLIMETERS);
#else
        private static readonly double MaxGrilleProjection = UnitUtils.ConvertToInternalUnits(1500, UnitTypeId.Millimeters);
        private static readonly double GrilleHalfWidth = UnitUtils.ConvertToInternalUnits(25, UnitTypeId.Millimeters);
#endif

        internal TagSpacesEquipmentAnalysis Analyze(Document document, ViewPlan view, Phase phase,
            Space[] spaces, PlanPoint[][] crop, TagSpacesEquipmentOptions options)
        {
            var result = new TagSpacesEquipmentAnalysis();
            var categories = new List<BuiltInCategory>();
            if (options.RadiatorType != null)
                categories.Add(BuiltInCategory.OST_MechanicalEquipment);
            if (options.GrilleType != null)
                categories.Add(BuiltInCategory.OST_DuctTerminal);

            // Коллектор ограничен конкретным видом, включая его категории, фильтры и варианты.
            FamilyInstance[] visible = new FilteredElementCollector(document, view.Id)
                .OfClass(typeof(FamilyInstance)).WherePasses(new ElementMulticategoryFilter(categories))
                .Cast<FamilyInstance>().Where(instance => IsVisible(instance, view, phase, crop)).ToArray();
            FamilyInstance[] radiators = visible.Where(instance => MatchesRule(instance, BuiltInCategory.OST_MechanicalEquipment)).ToArray();
            FamilyInstance[] grilles = visible.Where(instance => MatchesRule(instance, BuiltInCategory.OST_DuctTerminal)).ToArray();

            // Проверяем весь план до создания первой марки, даже если приборы находятся в разных пространствах.
            if (radiators.Length > 0 && grilles.Length > 0)
            {
                result.LayoutError = string.Format("План «{0}»: ошибка оформления. На одном виде видны радиаторы "
                    + "[id: {1}] и решётки [id: {2}]. Общие планы отопления и вентиляции запрещены. "
                    + "Исправьте видимость элементов и разделите планы по системам. План не маркирован.",
                    view.Name, Ids(radiators), Ids(grilles));
                return result;
            }

            // Геометрию вычисляем только для запасного поиска и переиспользуем между решётками.
            var spaceSolids = new Dictionary<ElementId, Solid>();
            using (var calculator = new SpatialElementGeometryCalculator(document))
            {
                foreach (FamilyInstance instance in radiators)
                    Associate(instance, spaces, phase, result.RadiatorSpaces, result.Messages, view.Name, null);
                foreach (FamilyInstance instance in grilles)
                    Associate(instance, spaces, phase, result.GrilleSpaces, result.Messages, view.Name,
                        () => FindSpaceBelow(instance, spaces, calculator, spaceSolids));
            }

            return result;
        }

        private static bool IsVisible(FamilyInstance instance, ViewPlan view, Phase phase, PlanPoint[][] crop)
        {
            if (instance.IsHidden(view) || view.GetCategoryHidden(instance.Category.Id))
                return false;
            if (view.IsTemporaryHideIsolateActive()
                && !view.IsElementVisibleInTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate, instance.Id))
                return false;

            ElementOnPhaseStatus status = instance.GetPhaseStatus(phase.Id);
            if (status != ElementOnPhaseStatus.Existing && status != ElementOnPhaseStatus.New)
                return false;
            if (crop.Length == 0)
                return true;

            // Коллектор Revit может вернуть элементы за обрезкой. Габарит используется только для видимости,
            // принадлежность пространству ниже определяется точкой в объёме, без пересечения габаритов.
            BoundingBoxXYZ bounds = instance.get_BoundingBox(view);
            if (bounds == null)
                return false;

            var projected = new List<PlanPoint>();
            foreach (double x in new[] { bounds.Min.X, bounds.Max.X })
                foreach (double y in new[] { bounds.Min.Y, bounds.Max.Y })
                    foreach (double z in new[] { bounds.Min.Z, bounds.Max.Z })
                    {
                        XYZ point = bounds.Transform.OfPoint(new XYZ(x, y, z));
                        projected.Add(new PlanPoint(point.DotProduct(view.RightDirection), point.DotProduct(view.UpDirection)));
                    }

            double minX = projected.Min(point => point.X);
            double maxX = projected.Max(point => point.X);
            double minY = projected.Min(point => point.Y);
            double maxY = projected.Max(point => point.Y);
            var rectangle = new[] { new PlanPoint(minX, minY), new PlanPoint(maxX, minY),
                new PlanPoint(maxX, maxY), new PlanPoint(minX, maxY) };
            return crop.Any(polygon => TagSpacesGeometry.Overlaps(polygon, rectangle));
        }

        private static void Associate(FamilyInstance instance, Space[] spaces, Phase phase,
            HashSet<ElementId> matches, List<string> messages, string viewName, Func<Space> findBelow)
        {
            string label = string.Format("План «{0}», прибор «{1}: {2}» [id: {3}]",
                viewName, instance.Symbol.FamilyName, instance.Name, instance.Id);
            try
            {
                XYZ point = findBelow != null ? (instance.Location as LocationPoint)?.Point
                    : instance.HasSpatialElementCalculationPoint ? instance.GetSpatialElementCalculationPoint()
                    : (instance.Location as LocationPoint)?.Point;
                if (point == null && findBelow == null)
                {
                    messages.Add(label + ": нет точки для определения пространства; проверьте принадлежность вручную.");
                    return;
                }

                Space native = instance.get_Space(phase);
                if (point != null && native != null && spaces.Any(space => space.Id.Equals(native.Id)) && native.IsPointInSpace(point))
                {
                    matches.Add(native.Id);
                    return;
                }

                Space[] containing = point == null ? new Space[0] : spaces.Where(space => space.Area > 0 && space.IsPointInSpace(point)).Take(2).ToArray();
                if (containing.Length == 1)
                {
                    matches.Add(containing[0].Id);
                    return;
                }

                if (containing.Length == 0 && findBelow != null)
                {
                    Space projected = findBelow();
                    if (projected != null)
                    {
                        matches.Add(projected.Id);
                        return;
                    }

                    messages.Add(label + ": пространство не найдено: LocationPoint, квадрат XY ±25 мм и проекция квадрата вниз до 1,5 м; специальная марка не назначена.");
                    return;
                }

                messages.Add(label + (containing.Length == 0
                        ? ": точка прибора вне объёма пространств плана; специальная марка не назначена."
                        : ": точка прибора попала в несколько пространств; исправьте пересечение пространств."));
            }
            catch (Autodesk.Revit.Exceptions.RegenerationFailedException)
            {
                throw;
            }
            catch (Exception ex)
            {
                messages.Add(label + ": не удалось определить пространство: " + ex.Message);
            }
        }

        private static Space FindSpaceBelow(FamilyInstance instance, Space[] spaces,
            SpatialElementGeometryCalculator calculator, Dictionary<ElementId, Solid> spaceSolids)
        {
            XYZ origin = (instance.Location as LocationPoint)?.Point;
            if (origin == null)
                return null;

            Space nearest = null;
            double nearestDistance = double.MaxValue;
            double half = GrilleHalfWidth;
            XYZ[] corners =
            {
                origin + new XYZ(-half, -half, 0), origin + new XYZ(half, -half, 0),
                origin + new XYZ(half, half, 0), origin + new XYZ(-half, half, 0)
            };

            using (var loop = new CurveLoop())
            {
                for (int i = 0; i < corners.Length; i++)
                    loop.Append(Line.CreateBound(corners[i], corners[(i + 1) % corners.Length]));

                // Верхняя грань — квадрат вокруг LocationPoint, объём под ней — третья ступень поиска.
                // Это точное пересечение объёмов, а не выборка точек и не габарит решётки.
                using (Solid search = GeometryCreationUtilities.CreateExtrusionGeometry(
                    new[] { loop }, -XYZ.BasisZ, MaxGrilleProjection))
                {
                    foreach (Space space in spaces)
                    {
                        if (space.Area <= 0)
                            continue;

                        if (!spaceSolids.TryGetValue(space.Id, out Solid solid))
                        {
                            solid = calculator.CalculateSpatialElementGeometry(space).GetGeometry();
                            spaceSolids.Add(space.Id, solid);
                        }

                        using (Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(search, solid, BooleanOperationsType.Intersect))
                        {
                            if (intersection.Volume <= 0)
                                continue;

                            // Габарит уже точного пересечения даёт верхнюю отметку первого попадания.
                            BoundingBoxXYZ bounds = intersection.GetBoundingBox();
                            double top = double.MinValue;
                            foreach (double x in new[] { bounds.Min.X, bounds.Max.X })
                                foreach (double y in new[] { bounds.Min.Y, bounds.Max.Y })
                                    foreach (double z in new[] { bounds.Min.Z, bounds.Max.Z })
                                        top = Math.Max(top, bounds.Transform.OfPoint(new XYZ(x, y, z)).Z);

                            double distance = Math.Max(0, origin.Z - top);
                            // Пересечение на исходной высоте (ступень 2) приоритетнее любого попадания ниже.
                            if (distance <= MaxGrilleProjection && distance < nearestDistance)
                            {
                                nearest = space;
                                nearestDistance = distance;
                            }
                        }
                    }
                }
            }

            return nearest;
        }

        private static bool MatchesRule(FamilyInstance instance, BuiltInCategory category)
            => instance.Category != null && instance.Category.Id.Equals(new ElementId(category))
                && TagSpacesEquipmentRules.MatchesFamilyName(category, instance.Symbol.FamilyName);

        private static string Ids(IEnumerable<FamilyInstance> instances)
            => string.Join(", ", instances.Take(10).Select(instance => instance.Id.ToString()));
    }
}
