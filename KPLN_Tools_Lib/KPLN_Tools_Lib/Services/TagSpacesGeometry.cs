using System;
using System.Collections.Generic;

namespace KPLN_Tools_Lib.Services
{
    internal struct PlanPoint
    {
        internal PlanPoint(double x, double y)
        {
            X = x;
            Y = y;
        }

        internal double X { get; }
        internal double Y { get; }
    }

    internal static class TagSpacesGeometry
    {
        internal static bool Overlaps(IList<PlanPoint> first, IList<PlanPoint> second)
        {
            foreach (PlanPoint point in first)
                if (Contains(second, point))
                    return true;
            foreach (PlanPoint point in second)
                if (Contains(first, point))
                    return true;

            for (int i = 0; i < first.Count; i++)
                for (int j = 0; j < second.Count; j++)
                {
                    PlanPoint a = first[i];
                    PlanPoint b = first[(i + 1) % first.Count];
                    PlanPoint c = second[j];
                    PlanPoint d = second[(j + 1) % second.Count];
                    if (Cross(a, b, c) * Cross(a, b, d) < 0 && Cross(c, d, a) * Cross(c, d, b) < 0)
                        return true;
                }

            return false;
        }

        private static double Cross(PlanPoint a, PlanPoint b, PlanPoint point)
            => (b.X - a.X) * (point.Y - a.Y) - (b.Y - a.Y) * (point.X - a.X);

        // Точка на границе включается; поддерживаются повёрнутые и невыпуклые контуры обрезки.
        internal static bool Contains(IList<PlanPoint> polygon, PlanPoint point)
        {
            const double tolerance = 1e-7;
            bool inside = false;
            for (int i = 0, j = polygon.Count - 1; i < polygon.Count; j = i++)
            {
                PlanPoint a = polygon[j];
                PlanPoint b = polygon[i];
                double dx = b.X - a.X;
                double dy = b.Y - a.Y;
                double length = Math.Sqrt(dx * dx + dy * dy);
                if (length <= tolerance)
                    continue;

                double cross = (point.X - a.X) * dy - (point.Y - a.Y) * dx;
                double dot = (point.X - a.X) * dx + (point.Y - a.Y) * dy;
                if (Math.Abs(cross) <= tolerance * length && dot >= -tolerance * length
                    && dot <= length * length + tolerance * length)
                    return true;

                if ((a.Y > point.Y) != (b.Y > point.Y)
                    && point.X < (b.X - a.X) * (point.Y - a.Y) / (b.Y - a.Y) + a.X)
                    inside = !inside;
            }

            return inside;
        }
    }
}
