using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    /// <summary>Объединение проекций в локальных координатах проёма, мм.</summary>
    internal static class ProjectionAreaUnion
    {
        internal const double ToleranceMm = 0.001;

        /// <summary>Только численный шум площади: 0,00001 процентного пункта.</summary>
        public static bool IsBelowRatio(double ratio, double threshold)
        {
            return ratio < threshold - 1e-7;
        }

        public static PathGeometry FromTriangles(IEnumerable<Point[]> triangles, bool fillEnclosedVoids)
        {
            StreamGeometry geometry = new StreamGeometry { FillRule = FillRule.Nonzero };
            using (StreamGeometryContext context = geometry.Open())
            {
                foreach (Point[] points in triangles)
                {
                    double cross = (points[1].X - points[0].X) * (points[2].Y - points[0].Y)
                        - (points[1].Y - points[0].Y) * (points[2].X - points[0].X);
                    if (Math.Abs(cross) < 1e-10)
                        continue;

                    // Одинаковый обход: наложенные треугольники не вычитают друг друга.
                    context.BeginFigure(points[0], true, true);
                    context.LineTo(points[cross > 0 ? 1 : 2], true, false);
                    context.LineTo(points[cross > 0 ? 2 : 1], true, false);
                }
            }

            PathGeometry outline = Combine(geometry, Geometry.Empty, GeometryCombineMode.Union);
            if (!fillEnclosedVoids)
                return outline;

            // Для ИОС занят весь наружный габарит, включая полость трубы/воздуховода.
            // Каждая замкнутая фигура заполняется отдельно; разрывы между фигурами сохраняются.
            PathGeometry filled = new PathGeometry();
            foreach (PathFigure figure in outline.GetFlattenedPathGeometry(ToleranceMm, ToleranceType.Absolute).Figures)
            {
                if (!figure.IsClosed)
                    throw new InvalidOperationException("Незамкнутая граница проекции ИОС");

                PathGeometry region = new PathGeometry(new[] { figure.Clone() });
                filled = Combine(filled, region, GeometryCombineMode.Union);
            }
            return filled;
        }

        public static PathGeometry Combine(Geometry first, Geometry second, GeometryCombineMode mode)
        {
            return Geometry.Combine(first, second, mode, null, ToleranceMm, ToleranceType.Absolute);
        }

        public static double Area(Geometry geometry)
        {
            double area = geometry.GetArea(ToleranceMm, ToleranceType.Absolute);
            if (double.IsNaN(area) || double.IsInfinity(area) || area < 0)
                throw new InvalidOperationException("Некорректная площадь проекции");
            return area;
        }
    }
}
