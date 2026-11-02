using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Media;
using Point = System.Windows.Point;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    /// <summary>Единая плоскость проёма и объединённая занятая проекция в пределах его глубины.</summary>
    internal sealed class HoleProjection
    {
        private const double _mmPerFoot = 304.8;
        private const double _cutRetryOffset = 0.5 / _mmPerFoot;
        private readonly XYZ _origin;
        private readonly XYZ _axis;
        private readonly XYZ _nearOrigin;
        private readonly XYZ _farOrigin;
        private readonly XYZ _right;
        private readonly XYZ _up;
        private readonly PathGeometry _aperture;
        private PathGeometry _occupied = new PathGeometry();

        public double OpeningArea { get; }
        public double OccupiedArea { get; private set; }

        public HoleProjection(FamilyInstance hole, Solid solid)
        {
            string familyName = hole.Symbol.FamilyName.ToLowerInvariant();
            Autodesk.Revit.DB.Transform transform = hole.GetTransform();
            // Проектное уточнение имеет приоритет над общими суффиксами семейства.
            bool isSetWallOpening = CheckHolesProjectRules.IsSetWallOpening(
                hole.Document.Title, hole.Symbol.FamilyName,
                hole.Category != null && hole.Category.Id.Equals(new ElementId(BuiltInCategory.OST_Windows)));

            XYZ axis;
            if (isSetWallOpening)
                axis = transform.BasisY.Normalize();
            else if (familyName.Contains("стена") || familyName.Contains("стене"))
                axis = transform.BasisY.Normalize();
            else if (familyName.Contains("перекрыт") || familyName.Contains("плита"))
                axis = transform.BasisZ.Normalize();
            else
                throw new InvalidOperationException("Не определено направление прохода для семейства отверстия " + hole.Symbol.FamilyName);

            _right = transform.BasisX.Normalize();
            _up = axis.CrossProduct(_right).Normalize();
            _axis = axis;
            _origin = solid.ComputeCentroid();

            // Подтверждаем постоянное сечение: две равные торцевые грани и площадь × глубина = объём.
            PlanarFace front = null;
            PlanarFace back = null;
            foreach (Face face in solid.Faces)
            {
                if (!(face is PlanarFace planar))
                    continue;
                double dot = planar.FaceNormal.DotProduct(axis);
                if (dot > 1 - 1e-8 && (front == null || planar.Area > front.Area))
                    front = planar;
                if (dot < -1 + 1e-8 && (back == null || planar.Area > back.Area))
                    back = planar;
            }
            if (front == null || back == null || front.Area <= 0)
                throw new InvalidOperationException("Не найдены торцевые грани проёма по оси семейства");

            _nearOrigin = back.Origin;
            _farOrigin = front.Origin;
            double depth = (_farOrigin - _nearOrigin).DotProduct(axis);
            double areaTolerance = Math.Max(front.Area * 1e-5, 1e-8);
            if (depth <= 1e-8 || Math.Abs(front.Area - back.Area) > areaTolerance)
                throw new InvalidOperationException("Геометрия отверстия не подтверждена как постоянное сечение по глубине");

            // Revit может приближённо вычислять Solid.Volume у криволинейных поверхностей.
            // Допуск объёма 0,1% разрешаем только после отдельного подтверждения постоянства
            // боковых граней. Для остальных тел сохраняем исходный строгий допуск 0,001%.
            bool hasConstantCurvedSection = HasConstantCurvedSection(solid);
            double volumeAreaTolerance = hasConstantCurvedSection
                ? Math.Max(front.Area * 0.001, 1e-8)
                : areaTolerance;
            if (Math.Abs(Math.Abs(solid.Volume) / depth - front.Area) > volumeAreaTolerance)
                throw new InvalidOperationException("Геометрия отверстия не подтверждена как постоянное сечение по глубине");

            // Сетка Triangulate даже на уровне 1.0 недостаточно точно описывает малые круги.
            // Для подтверждённого постоянного сечения берём контур торца по его дугам.
            _aperture = (hasConstantCurvedSection ? TryProjectEndFace(front) : null) ?? Project(solid, false);
            double projectedArea = ProjectionAreaUnion.Area(_aperture) / (_mmPerFoot * _mmPerFoot);
            // Независимо от способа построения не выдаём слишком грубый контур за точный.
            if (Math.Abs(projectedArea - front.Area) > Math.Max(front.Area * 0.001, 1e-8))
                throw new InvalidOperationException("Проекция проёма не совпадает с площадью торца в пределах 0,1%");
            OpeningArea = projectedArea;
        }

        /// <summary>Подтверждает постоянное криволинейное сечение по аналитическим поверхностям.</summary>
        private bool HasConstantCurvedSection(Solid solid)
        {
            bool hasCylinder = false;
            foreach (Face face in solid.Faces)
            {
                if (face is PlanarFace planar)
                {
                    double dot = planar.FaceNormal.DotProduct(_axis);
                    if (Math.Abs(dot) <= 1e-8)
                        continue;

                    // Допускаются только боковые плоскости вдоль оси и торцы на границах
                    // глубины. Промежуточная ступень или наклонная грань не проходят проверку.
                    if (Math.Abs(dot) < 1 - 1e-8)
                        return false;
                    XYZ endOrigin = dot > 0 ? _farOrigin : _nearOrigin;
                    if (Math.Abs((planar.Origin - endOrigin).DotProduct(_axis)) > 1e-7)
                        return false;
                }
                else if (face is CylindricalFace cylinder)
                {
                    if (Math.Abs(cylinder.Axis.Normalize().DotProduct(_axis)) < 1 - 1e-8)
                        return false;
                    hasCylinder = true;
                }
                else
                    return false;
            }

            return hasCylinder;
        }

        /// <summary>Контур торца из отрезков и дуг с погрешностью хорды не более 0,001 мм.</summary>
        private PathGeometry TryProjectEndFace(PlanarFace face)
        {
            PathGeometry geometry = new PathGeometry { FillRule = FillRule.EvenOdd };
            int pointCount = 0;
            foreach (CurveLoop loop in face.GetEdgesAsCurveLoops())
            {
                PathFigure figure = new PathFigure { IsClosed = true, IsFilled = true };
                bool started = false;
                foreach (Curve curve in loop)
                {
                    if (!(curve is Line) && !(curve is Arc))
                        return null;

                    double start = curve.IsBound ? curve.GetEndParameter(0) : 0;
                    double end = curve.IsBound ? curve.GetEndParameter(1) : 2 * Math.PI;
                    if (!started)
                    {
                        figure.StartPoint = ProjectPoint(curve.Evaluate(start, false));
                        started = true;
                    }

                    int segments = 1;
                    if (curve is Arc arc)
                    {
                        double radiusMm = arc.Radius * _mmPerFoot;
                        double angleStep = Math.Min(Math.PI / 18,
                            2 * Math.Acos(Math.Max(-1, 1 - ProjectionAreaUnion.ToleranceMm / radiusMm)));
                        double required = Math.Ceiling(Math.Abs(end - start) / angleStep);
                        if (double.IsNaN(required) || double.IsInfinity(required) || required > 100000)
                            throw new InvalidOperationException("Слишком сложный контур торца отверстия");
                        segments = Math.Max(1, (int)required);
                    }

                    pointCount += segments;
                    if (pointCount > 100000)
                        throw new InvalidOperationException("Слишком сложный контур торца отверстия");
                    for (int i = 1; i <= segments; i++)
                    {
                        XYZ point = curve.Evaluate(start + (end - start) * i / segments, false);
                        figure.Segments.Add(new System.Windows.Media.LineSegment(ProjectPoint(point), true));
                    }
                }

                if (!started)
                    throw new InvalidOperationException("Пустой контур торца отверстия");
                geometry.Figures.Add(figure);
            }
            return geometry;
        }

        private Point ProjectPoint(XYZ point)
        {
            XYZ vector = point - _origin;
            return new Point(vector.DotProduct(_right) * _mmPerFoot, vector.DotProduct(_up) * _mmPerFoot);
        }

        /// <summary>Оставляет только участок тела в пределах глубины проёма.</summary>
        public Solid ClipToDepth(Solid solid)
        {
            // У одного элемента могут быть отдельные тела вне проёма. Их отсутствие пересечения
            // не является сбоем Revit и не должно отменять анализ остальных тел этого элемента.
            if (IsOutsideDepth(solid))
                return null;

            using (Solid clipped = CutChecked(solid, _axis, _nearOrigin))
            {
                if (clipped == null)
                    return null;

                return CutChecked(clipped, _axis.Negate(), _farOrigin);
            }
        }

        /// <summary>Консервативный отбор тела по глубине, включая допуск единственного повтора.</summary>
        private bool IsOutsideDepth(Solid solid)
        {
            using (BoundingBoxXYZ bounds = solid.GetBoundingBox())
            {
                if (bounds == null)
                    return false;

                double min = double.MaxValue;
                double max = double.MinValue;
                Autodesk.Revit.DB.Transform transform = bounds.Transform;
                // Проверяем все восемь углов в координатах модели: локальные Min/Max
                // не дают диапазон вдоль оси отверстия для повёрнутого тела или связи.
                for (int i = 0; i < 8; i++)
                {
                    XYZ corner = transform.OfPoint(new XYZ(
                        (i & 1) == 0 ? bounds.Min.X : bounds.Max.X,
                        (i & 2) == 0 ? bounds.Min.Y : bounds.Max.Y,
                        (i & 4) == 0 ? bounds.Min.Z : bounds.Max.Z));
                    double distance = (corner - _nearOrigin).DotProduct(_axis);
                    min = Math.Min(min, distance);
                    max = Math.Max(max, distance);
                }

                double depth = (_farOrigin - _nearOrigin).DotProduct(_axis);
                const double tolerance = 1e-9;
                return max < -_cutRetryOffset - tolerance
                    || min > depth + _cutRetryOffset + tolerance;
            }
        }

        /// <summary>При null повторяем отсечение один раз, отступив наружу на 0,5 мм.</summary>
        private static Solid CutChecked(Solid solid, XYZ normal, XYZ origin)
        {
            Solid result;
            using (Plane plane = Plane.CreateByNormalAndOrigin(normal, origin))
                result = BooleanOperationsUtils.CutWithHalfSpace(solid, plane);

            if (result == null)
            {
                // Нормаль направлена в сохраняемую область: отступ против неё расширяет глубину.
                // Повторяем на исходном теле, без накопления смещения и дополнительных попыток.
                XYZ shiftedOrigin = origin - normal.Normalize() * _cutRetryOffset;
                try
                {
                    using (Plane shiftedPlane = Plane.CreateByNormalAndOrigin(normal, shiftedOrigin))
                        result = BooleanOperationsUtils.CutWithHalfSpace(solid, shiftedPlane);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "Не удалось обрезать геометрию по глубине проёма после единственного смещения на 0,5 мм. " +
                        "Проверка неполная: " + ex.Message, ex);
                }

                if (result == null)
                    throw new InvalidOperationException(
                        "Revit повторно вернул null при обрезке по глубине проёма после смещения на 0,5 мм. " +
                        "Проверка неполная");
            }

            // Не выдаём тело без объёма за подтверждённое отсутствие пересечения.
            try
            {
                if (result.Faces.Size == 0 || Math.Abs(result.Volume) <= 1e-12)
                    throw new InvalidOperationException(
                        "Revit вернул тело без объёма при обрезке по глубине проёма. Проверка неполная");

                return result;
            }
            catch
            {
                result.Dispose();
                throw;
            }
        }

        /// <summary>Добавляет тела одного элемента, обрезанные по глубине проёма, атомарно.</summary>
        public bool Add(IList<Solid> intersections)
        {
            PathGeometry elementProjection = new PathGeometry();
            foreach (Solid intersection in intersections)
                elementProjection = ProjectionAreaUnion.Combine(elementProjection, Project(intersection, true), GeometryCombineMode.Union);

            // Контуром проёма обрезаем после заполнения полостей ИОС: иначе срезанная изоляция
            // превращается в незамкнутую оболочку, а её внутренняя область ошибочно считается свободной.
            elementProjection = ProjectionAreaUnion.Combine(elementProjection, _aperture, GeometryCombineMode.Intersect);
            if (ProjectionAreaUnion.Area(elementProjection) <= 1e-8)
                return false;

            PathGeometry combined = ProjectionAreaUnion.Combine(_occupied, elementProjection, GeometryCombineMode.Union);
            double occupiedArea = ProjectionAreaUnion.Area(combined) / (_mmPerFoot * _mmPerFoot);
            if (occupiedArea > OpeningArea + Math.Max(OpeningArea * 1e-6, 1e-8))
                throw new InvalidOperationException("Объединённая площадь вышла за пределы проёма");

            _occupied = combined;
            OccupiedArea = occupiedArea;
            return true;
        }

        private PathGeometry Project(Solid solid, bool fillEnclosedVoids)
        {
            List<Point[]> triangles = new List<Point[]>();
            foreach (Face face in solid.Faces)
            {
                Mesh mesh = face.Triangulate(1.0);
                if (triangles.Count + mesh.NumTriangles > 100000)
                    throw new InvalidOperationException("Слишком сложная геометрия для проекции (более 100000 треугольников)");

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    Point[] points = new Point[3];
                    for (int j = 0; j < 3; j++)
                    {
                        XYZ vector = triangle.get_Vertex(j) - _origin;
                        points[j] = new Point(vector.DotProduct(_right) * _mmPerFoot, vector.DotProduct(_up) * _mmPerFoot);
                    }
                    triangles.Add(points);
                }
            }
            if (triangles.Count == 0)
                throw new InvalidOperationException("Нет поверхности для проекции");
            return ProjectionAreaUnion.FromTriangles(triangles, fillEnclosedVoids);
        }
    }
}
