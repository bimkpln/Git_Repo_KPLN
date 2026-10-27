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
            XYZ axis;
            if (familyName.Contains("стена") || familyName.Contains("стене"))
                axis = transform.BasisY.Normalize();
            else if (familyName.Contains("перекрыт") || familyName.Contains("плита"))
                axis = transform.BasisZ.Normalize();
            else
                throw new InvalidOperationException("Не определено направление прохода для семейства отверстия " + hole.Symbol.FamilyName);

            _right = transform.BasisX.Normalize();
            _up = axis.CrossProduct(_right).Normalize();
            _axis = axis;
            _origin = solid.ComputeCentroid();

            // Подтверждаем призматический проём: две равные торцевые грани и площадь × глубина = объём.
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
            if (depth <= 1e-8 || Math.Abs(front.Area - back.Area) > areaTolerance
                || Math.Abs(Math.Abs(solid.Volume) / depth - front.Area) > areaTolerance)
                throw new InvalidOperationException("Геометрия отверстия не подтверждена как постоянное сечение по глубине");

            _aperture = Project(solid, false);
            double projectedArea = ProjectionAreaUnion.Area(_aperture) / (_mmPerFoot * _mmPerFoot);
            // Кривые границы аппроксимируются сеткой Revit. Слишком грубый контур не выдаём за точный.
            if (Math.Abs(projectedArea - front.Area) > Math.Max(front.Area * 0.001, 1e-8))
                throw new InvalidOperationException("Проекция проёма не совпадает с площадью торца в пределах 0,1%");
            OpeningArea = projectedArea;
        }

        /// <summary>Оставляет только участок тела в пределах глубины проёма.</summary>
        public Solid ClipToDepth(Solid solid)
        {
            using (Plane near = Plane.CreateByNormalAndOrigin(_axis, _nearOrigin))
            using (Solid clipped = BooleanOperationsUtils.CutWithHalfSpace(solid, near))
            {
                if (clipped == null || clipped.Faces.Size == 0 || Math.Abs(clipped.Volume) <= 1e-12)
                    return null;

                using (Plane far = Plane.CreateByNormalAndOrigin(_axis.Negate(), _farOrigin))
                    return BooleanOperationsUtils.CutWithHalfSpace(clipped, far);
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
