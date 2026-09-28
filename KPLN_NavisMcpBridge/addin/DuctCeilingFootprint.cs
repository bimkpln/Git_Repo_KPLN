using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_NavisMcpBridge
{
    public sealed class ClashFootprintDto
    {
        public bool Usable;
        public string Reason;
        public string Source = "duct-ceiling-mesh-section";
        public int DuctSide;
        public string ServiceKind;
        public string NominalSize;
        public string NominalSizeSource;
        public double? NominalWidth;
        public double? NominalHeight;
        public double? NominalDiameter;
        public double? InsulationThickness;
        public Point3Dto Normal;
        public Point3Dto PlaneOrigin;
        public double MinEdge;
        public double MaxEdge;
        public double Area;
        public double DuctSliceArea;
        public List<FootprintFragmentDiagnostics> DuctMeshDiagnostics;
        public List<FootprintFragmentDiagnostics> CeilingMeshDiagnostics;
    }

    internal static partial class DuctCeilingFootprint
    {
        private const double Eps = 1e-7;
        private struct P2 { public double X, Y; public P2(double x, double y) { X = x; Y = y; } }
        private sealed class Face { public Point3Dto A, B, C, N; public double Area; }
        private sealed class Direction { public Point3Dto N; public double Area; }

        internal static ClashFootprintDto Failure(string reason) => new ClashFootprintDto { Reason = reason };

        internal static ClashFootprintDto Measure(IList<Point3Dto> duct, IList<Point3Dto> ceiling, Point3Dto contact)
        {
            if (!Valid(duct) || !Valid(ceiling) || !Finite(contact)) return Failure("missing-or-invalid-mesh");
            var faces = Faces(ceiling);
            if (faces.Count == 0) return Failure("no-ceiling-faces");
            var directions = new List<Direction>();
            foreach (var f in faces)
            {
                var group = directions.FirstOrDefault(d => Math.Abs(Dot(d.N, f.N)) > 1 - 1e-6);
                if (group == null) { group = new Direction { N = f.N }; directions.Add(group); }
                group.Area += f.Area;
            }
            directions = directions.OrderByDescending(d => d.Area).ToList();
            if (directions.Count > 1 && directions[0].Area < 2 * directions[1].Area)
                return Failure("ambiguous-ceiling-plane");
            var normal = directions[0].N;
            var origin = contact;
            var u = Unit(Cross(normal, Math.Abs(normal.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
            var v = Cross(normal, u);
            var broad = faces.Where(f => Math.Abs(Dot(f.N, normal)) > 1 - 1e-6).ToList();
            var offsets = new List<double>();
            foreach (var f in broad)
            {
                var d = Dot(Sub(f.A, origin), normal);
                if (!offsets.Any(x => Math.Abs(x - d) < Eps)) offsets.Add(d);
            }
            // The contact-nearest broad surface with actual overlap defines this collision footprint.
            foreach (var offset in offsets.OrderBy(Math.Abs))
            {
                var planeOrigin = Add(origin, normal, offset);
                var slicePoints = new List<P2>();
                for (int i = 0; i < duct.Count; i += 3)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        var a = duct[i + j]; var b = duct[i + (j + 1) % 3];
                        var da = Dot(Sub(a, planeOrigin), normal); var db = Dot(Sub(b, planeOrigin), normal);
                        if (Math.Abs(da) <= Eps) slicePoints.Add(Project(a, planeOrigin, u, v));
                        if ((da < -Eps && db > Eps) || (da > Eps && db < -Eps))
                            slicePoints.Add(Project(Add(a, Sub(b, a), da / (da - db)), planeOrigin, u, v));
                    }
                }
                var slice = Hull(slicePoints);
                if (slice.Count < 3 || Area(slice) < Eps * Eps) continue;
                var sliceArea = Area(slice);
                var overlapPoints = new List<P2>();
                double overlapArea = 0;
                var seen = new HashSet<string>();
                foreach (var face in broad)
                {
                    if (Math.Abs(Dot(Sub(face.A, origin), normal) - offset) > Eps) continue;
                    var triangle = new List<P2> { Project(face.A, planeOrigin, u, v),
                        Project(face.B, planeOrigin, u, v), Project(face.C, planeOrigin, u, v) };
                    // Ignore duplicated triangles (e.g. two-sided export), not holes or separated patches.
                    var key = string.Join(";", triangle.Select(p => Math.Round(p.X / Eps) + "," + Math.Round(p.Y / Eps)).OrderBy(s => s, StringComparer.Ordinal));
                    if (!seen.Add(key)) continue;
                    if (SignedArea(triangle) < 0) triangle.Reverse();
                    var clipped = Clip(slice, triangle);
                    var area = Area(clipped);
                    if (area <= Eps * Eps) continue;
                    overlapArea += area;
                    overlapPoints.AddRange(clipped);
                }
                if (overlapArea <= Eps * Eps) continue;
                if (overlapArea > sliceArea * (1 + 1e-5)) return Failure("overlapping-ceiling-faces");
                var hull = Hull(overlapPoints);
                if (hull.Count < 3) continue;
                var sides = Rectangle(hull);
                return new ClashFootprintDto { Usable = true, Normal = normal, PlaneOrigin = planeOrigin,
                    MinEdge = sides[0], MaxEdge = sides[1], Area = overlapArea, DuctSliceArea = sliceArea };
            }
            return Failure("no-broad-face-intersection");
        }

        private static List<P2> Clip(List<P2> polygon, List<P2> triangle)
        {
            var output = polygon.ToList();
            for (int i = 0; i < 3 && output.Count > 0; i++)
            {
                var a = triangle[i]; var b = triangle[(i + 1) % 3];
                var input = output; output = new List<P2>();
                var previous = input[input.Count - 1]; var pd = Cross2(a, b, previous);
                foreach (var current in input)
                {
                    var cd = Cross2(a, b, current);
                    if ((pd >= -Eps) != (cd >= -Eps))
                    {
                        var t = pd / (pd - cd);
                        output.Add(new P2(previous.X + t * (current.X - previous.X), previous.Y + t * (current.Y - previous.Y)));
                    }
                    if (cd >= -Eps) output.Add(current);
                    previous = current; pd = cd;
                }
            }
            return output;
        }

        private static List<P2> Hull(List<P2> points)
        {
            var sorted = points.OrderBy(p => p.X).ThenBy(p => p.Y).ToList();
            var unique = new List<P2>();
            foreach (var p in sorted)
                if (unique.Count == 0 || Math.Abs(p.X - unique[unique.Count - 1].X) > Eps ||
                    Math.Abs(p.Y - unique[unique.Count - 1].Y) > Eps) unique.Add(p);
            if (unique.Count <= 2) return unique;
            var lower = new List<P2>(); var upper = new List<P2>();
            foreach (var p in unique)
            {
                while (lower.Count >= 2 && Cross2(lower[lower.Count - 2], lower[lower.Count - 1], p) <= Eps * Eps) lower.RemoveAt(lower.Count - 1);
                lower.Add(p);
            }
            for (int i = unique.Count - 1; i >= 0; i--)
            {
                var p = unique[i];
                while (upper.Count >= 2 && Cross2(upper[upper.Count - 2], upper[upper.Count - 1], p) <= Eps * Eps) upper.RemoveAt(upper.Count - 1);
                upper.Add(p);
            }
            lower.RemoveAt(lower.Count - 1); upper.RemoveAt(upper.Count - 1); lower.AddRange(upper); return lower;
        }

        private static double[] Rectangle(List<P2> hull)
        {
            double best = double.PositiveInfinity, first = 0, second = 0;
            for (int i = 0; i < hull.Count; i++)
            {
                var a = hull[i]; var b = hull[(i + 1) % hull.Count];
                var dx = b.X - a.X; var dy = b.Y - a.Y; var n = Math.Sqrt(dx * dx + dy * dy);
                if (n <= Eps) continue;
                dx /= n; dy /= n;
                var along = hull.Select(p => p.X * dx + p.Y * dy).ToArray();
                var across = hull.Select(p => -p.X * dy + p.Y * dx).ToArray();
                var x = along.Max() - along.Min(); var y = across.Max() - across.Min();
                if (x * y < best) { best = x * y; first = Math.Min(x, y); second = Math.Max(x, y); }
            }
            return new[] { first, second };
        }

        private static List<Face> Faces(IList<Point3Dto> mesh)
        {
            var faces = new List<Face>();
            for (int i = 0; i < mesh.Count; i += 3)
            {
                var a = mesh[i]; var b = mesh[i + 1]; var c = mesh[i + 2]; var n = Cross(Sub(b, a), Sub(c, a));
                var length = Math.Sqrt(Dot(n, n));
                if (length > Eps * Eps) faces.Add(new Face { A = a, B = b, C = c, N = P(n.X / length, n.Y / length, n.Z / length), Area = length * .5 });
            }
            return faces;
        }
        private static bool Valid(IList<Point3Dto> mesh) => mesh != null && mesh.Count >= 6 && mesh.Count % 3 == 0 && mesh.All(Finite);
        private static bool Finite(Point3Dto p) => p != null && new[] { p.X, p.Y, p.Z }.All(x => !double.IsNaN(x) && !double.IsInfinity(x));
        private static Point3Dto P(double x, double y, double z) => new Point3Dto { X = x, Y = y, Z = z };
        private static Point3Dto Sub(Point3Dto a, Point3Dto b) => P(a.X - b.X, a.Y - b.Y, a.Z - b.Z);
        private static Point3Dto Add(Point3Dto a, Point3Dto b, double t) => P(a.X + t * b.X, a.Y + t * b.Y, a.Z + t * b.Z);
        private static double Dot(Point3Dto a, Point3Dto b) => a.X * b.X + a.Y * b.Y + a.Z * b.Z;
        private static Point3Dto Cross(Point3Dto a, Point3Dto b) => P(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
        private static Point3Dto Unit(Point3Dto p) { var n = Math.Sqrt(Dot(p, p)); return P(p.X / n, p.Y / n, p.Z / n); }
        private static P2 Project(Point3Dto p, Point3Dto origin, Point3Dto u, Point3Dto v) { var d = Sub(p, origin); return new P2(Dot(d, u), Dot(d, v)); }
        private static double Cross2(P2 a, P2 b, P2 c) => (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X);
        private static double SignedArea(List<P2> p) { double sum = 0; for (int i = 0; i < p.Count; i++) sum += p[i].X * p[(i + 1) % p.Count].Y - p[i].Y * p[(i + 1) % p.Count].X; return sum * .5; }
        private static double Area(List<P2> p) => Math.Abs(SignedArea(p));
    }
}
