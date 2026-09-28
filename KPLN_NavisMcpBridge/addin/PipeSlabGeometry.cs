using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_NavisMcpBridge
{
    public sealed class PipeSlabSectionDto
    {
        public Point3Dto Origin;
        public List<Point3Dto> Contour;
        public double Area, CoveredArea;
    }

    public sealed class PipeSlabDto
    {
        public string Source = "pipe-slab-mesh-1";
        public bool Usable;
        public string Reason;
        public int PipeSide;
        public string PipeId, SlabId;
        public Point3Dto PipeAxis, AxisOrigin, SlabNormal;
        public double OuterRadius, PipeLength, SlabThickness;
        public bool EdgeContact, FullSectionCovered;
        public List<PipeSlabSectionDto> Sections = new List<PipeSlabSectionDto>();
    }

    // Shares the tested triangle-section/clipping kernel with duct footprints.
    // No opening thresholds or Navisworks status decisions belong in this layer.
    internal static partial class DuctCeilingFootprint
    {
        private sealed class Cylinder
        {
            public Point3Dto Axis, Origin;
            public double Radius, Low, High;
        }

        internal static PipeSlabDto MeasurePipeSlab(IList<Point3Dto> pipe, IList<Point3Dto> slab, Point3Dto contact)
        {
            var result = new PipeSlabDto();
            if (!Valid(pipe) || !Valid(slab) || !Finite(contact)) { result.Reason = "missing-or-invalid-mesh"; return result; }
            var cylinder = FitCylinder(pipe, out var cylinderReason);
            if (cylinder == null) { result.Reason = cylinderReason; return result; }
            result.PipeAxis = cylinder.Axis;
            result.AxisOrigin = cylinder.Origin;
            result.OuterRadius = cylinder.Radius;
            result.PipeLength = cylinder.High - cylinder.Low;
            var faces = UniqueFaces(slab);
            if (!ClosedMesh(faces)) { result.Reason = "slab-mesh-not-closed"; return result; }
            var directions = new List<Direction>();
            foreach (var f in faces)
            {
                var d = directions.FirstOrDefault(x => Math.Abs(Dot(x.N, f.N)) > 1 - 1e-6);
                if (d == null) { d = new Direction { N = f.N }; directions.Add(d); }
                d.Area += f.Area;
            }
            directions = directions.OrderByDescending(x => x.Area).ToList();
            if (directions.Count < 2 || directions[0].Area < 2 * directions[1].Area)
            { result.Reason = "ambiguous-slab-normal"; return result; }
            var normal = directions[0].N;
            result.SlabNormal = normal;
            var broad = faces.Where(f => Math.Abs(Dot(f.N, normal)) > 1 - 1e-6).ToList();
            var levels = new List<double>();
            foreach (var f in broad)
            {
                var level = Dot(Sub(f.A, contact), normal);
                if (!levels.Any(x => Math.Abs(x - level) < Eps)) levels.Add(level);
            }
            // Stepped/multi-layer slabs require a local paired-surface implementation.
            if (levels.Count != 2) { result.Reason = "slab-does-not-have-two-parallel-broad-planes"; return result; }
            result.SlabThickness = levels.Max() - levels.Min();
            foreach (var face in faces.Except(broad))
                if (CylinderTouchesTriangle(cylinder, face)) { result.EdgeContact = true; break; }
            var u = Unit(Cross(normal, Math.Abs(normal.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
            var v = Cross(normal, u);
            foreach (var level in levels.OrderBy(Math.Abs))
            {
                var origin = Add(contact, normal, level);
                var slice = SliceMesh(pipe, origin, normal, u, v);
                if (slice.Count < 3 || Area(slice) < Eps * Eps) continue;
                double covered = 0;
                foreach (var face in broad)
                {
                    if (Math.Abs(Dot(Sub(face.A, contact), normal) - level) > Eps) continue;
                    var tri = new List<P2> { Project(face.A, origin, u, v), Project(face.B, origin, u, v), Project(face.C, origin, u, v) };
                    if (SignedArea(tri) < 0) tri.Reverse();
                    covered += Area(Clip(slice, tri));
                }
                var area = Area(slice);
                if (covered > area * (1 + 1e-5)) { result.Reason = "overlapping-slab-faces"; return result; }
                result.Sections.Add(new PipeSlabSectionDto { Origin = origin, Area = area, CoveredArea = covered,
                    Contour = slice.Select(p => Add(Add(origin, u, p.X), v, p.Y)).ToList() });
            }
            result.FullSectionCovered = result.Sections.Count > 0 &&
                result.Sections.All(s => s.CoveredArea >= s.Area * (1 - 1e-5));
            result.Usable = true;
            return result;
        }

        private static List<P2> SliceMesh(IList<Point3Dto> mesh, Point3Dto origin, Point3Dto n, Point3Dto u, Point3Dto v)
        {
            var points = new List<P2>();
            for (int i = 0; i < mesh.Count; i += 3)
                for (int j = 0; j < 3; j++)
                {
                    var a = mesh[i + j]; var b = mesh[i + (j + 1) % 3];
                    var da = Dot(Sub(a, origin), n); var db = Dot(Sub(b, origin), n);
                    if (Math.Abs(da) <= Eps) points.Add(Project(a, origin, u, v));
                    if ((da < -Eps && db > Eps) || (da > Eps && db < -Eps))
                        points.Add(Project(Add(a, Sub(b, a), da / (da - db)), origin, u, v));
                }
            return Hull(points);
        }

        private static Cylinder FitCylinder(IList<Point3Dto> mesh, out string reason)
        {
            reason = "not-a-verified-straight-circular-pipe";
            var faces = UniqueFaces(mesh);
            var normals = new List<Point3Dto>();
            foreach (var f in faces)
                if (!normals.Any(n => Math.Abs(Dot(n, f.N)) > 1 - 1e-6)) normals.Add(f.N);
            if (normals.Count < 5 || normals.Count > 256) { reason = "pipe-normal-count-outside-cylinder-range"; return null; }
            var candidates = normals.ToList();
            // Cap normals work for very short pipes; crosses of lateral normals also
            // recover uncapped pipes. No longest-box or largest-PCA-axis assumption.
            for (int i = 0; i < normals.Count; i++)
                for (int j = i + 1; j < normals.Count; j++)
                {
                    var c = Cross(normals[i], normals[j]);
                    if (Dot(c, c) < 1e-4) continue;
                    c = Unit(c);
                    if (!candidates.Any(n => Math.Abs(Dot(n, c)) > 1 - 1e-6)) candidates.Add(c);
                    if (candidates.Count > 512) { reason = "pipe-axis-candidates-ambiguous"; return null; }
                }
            Cylinder found = null;
            foreach (var axis in candidates)
            {
                if (faces.Any(f => Math.Abs(Dot(axis, f.N)) > 1e-5 && Math.Abs(Dot(axis, f.N)) < 1 - 1e-5)) continue;
                var origin = mesh[0];
                var u = Unit(Cross(axis, Math.Abs(axis.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
                var v = Cross(axis, u);
                var hull = Hull(mesh.Select(p => Project(p, origin, u, v)).ToList());
                if (hull.Count < 8) continue;
                // Circumcenter from well-separated hull vertices, then validate ALL
                // outer vertices and lateral faces. World coordinates are localized.
                var a = hull[0]; var b = hull[hull.Count / 3]; var c = hull[2 * hull.Count / 3];
                var det = 2 * (a.X * (b.Y - c.Y) + b.X * (c.Y - a.Y) + c.X * (a.Y - b.Y));
                if (Math.Abs(det) < Eps * Eps) continue;
                var aa = a.X * a.X + a.Y * a.Y; var bb = b.X * b.X + b.Y * b.Y; var cc = c.X * c.X + c.Y * c.Y;
                var center = new P2((aa * (b.Y - c.Y) + bb * (c.Y - a.Y) + cc * (a.Y - b.Y)) / det,
                    (aa * (c.X - b.X) + bb * (a.X - c.X) + cc * (b.X - a.X)) / det);
                var radius = Math.Sqrt((a.X - center.X) * (a.X - center.X) + (a.Y - center.Y) * (a.Y - center.Y));
                var tolerance = Math.Max(1e-6, radius * .001);
                if (radius <= Eps || hull.Any(p => Math.Abs(Radial(p, center) - radius) > tolerance)) continue;
                var angles = hull.Select(p => Math.Atan2(p.Y - center.Y, p.X - center.X)).OrderBy(x => x).ToList();
                angles.Add(angles[0] + 2 * Math.PI);
                if (Enumerable.Range(0, angles.Count - 1).Any(i => angles[i + 1] - angles[i] > Math.PI / 4 + 1e-5)) continue;
                var lateral = faces.Where(f => Math.Abs(Dot(axis, f.N)) < 1e-5).ToList();
                if (lateral.Count < 16) continue;
                var radii = lateral.SelectMany(f => new[] { f.A, f.B, f.C }).Select(p => Radial(Project(p, origin, u, v), center)).ToList();
                var inner = radii.Min();
                if (inner <= Eps || radii.Any(r => Math.Abs(r - radius) > tolerance && Math.Abs(r - inner) > tolerance)) continue;
                var axisOrigin = Add(Add(origin, u, center.X), v, center.Y);
                var axial = mesh.Select(p => Dot(Sub(p, axisOrigin), axis)).ToList();
                if (axial.Max() - axial.Min() <= Eps) continue;
                var outerArea = lateral.Where(f => new[] { f.A, f.B, f.C }.All(p =>
                    Math.Abs(Radial(Project(p, origin, u, v), center) - radius) <= tolerance)).Sum(f => f.Area);
                var perimeter = Enumerable.Range(0, hull.Count).Sum(i => Radial(hull[i], hull[(i + 1) % hull.Count]));
                if (Math.Abs(outerArea - perimeter * (axial.Max() - axial.Min())) > perimeter * (axial.Max() - axial.Min()) * 1e-5) continue;
                // Navisworks may tessellate one straight pipe with intermediate
                // axial rings. The complete outer lateral-area check above proves
                // continuous coverage of the full axial interval and still rejects
                // a gap between disconnected coaxial pieces; intermediate vertices
                // themselves are therefore valid cylinder evidence.
                if (found != null && Math.Abs(Dot(found.Axis, axis)) < 1 - 1e-6) return null;
                found = new Cylinder { Axis = axis, Origin = axisOrigin, Radius = radius, Low = axial.Min(), High = axial.Max() };
            }
            if (found == null) reason = "pipe-mesh-does-not-form-one-complete-cylinder";
            return found;
        }

        private static double Radial(P2 p, P2 c) => Math.Sqrt((p.X - c.X) * (p.X - c.X) + (p.Y - c.Y) * (p.Y - c.Y));
        private static string PointKey(Point3Dto p) => Math.Round(p.X / 1e-6) + "," + Math.Round(p.Y / 1e-6) + "," + Math.Round(p.Z / 1e-6);
        private static List<Face> UniqueFaces(IList<Point3Dto> mesh)
        {
            var seen = new HashSet<string>();
            return Faces(mesh).Where(f => seen.Add(string.Join(";", new[] { PointKey(f.A), PointKey(f.B), PointKey(f.C) }.OrderBy(x => x, StringComparer.Ordinal)))).ToList();
        }
        private static bool ClosedMesh(List<Face> faces)
        {
            var counts = new Dictionary<string, int>();
            foreach (var f in faces)
            {
                var p = new[] { PointKey(f.A), PointKey(f.B), PointKey(f.C) };
                for (int i = 0; i < 3; i++)
                {
                    var k = string.Join(";", new[] { p[i], p[(i + 1) % 3] }.OrderBy(x => x, StringComparer.Ordinal));
                    counts[k] = counts.TryGetValue(k, out var count) ? count + 1 : 1;
                }
            }
            return counts.Count > 0 && counts.Values.All(c => c == 2);
        }

        private static bool CylinderTouchesTriangle(Cylinder cylinder, Face face)
        {
            var polygon = new List<Point3Dto> { face.A, face.B, face.C };
            polygon = ClipAxial(polygon, cylinder, cylinder.Low, 1);
            polygon = ClipAxial(polygon, cylinder, cylinder.High, -1);
            if (polygon.Count == 0) return false;
            var u = Unit(Cross(cylinder.Axis, Math.Abs(cylinder.Axis.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
            var v = Cross(cylinder.Axis, u);
            var points = Hull(polygon.Select(p => Project(p, cylinder.Origin, u, v)).ToList());
            var zero = new P2(0, 0);
            if (points.Count >= 3 && Enumerable.Range(0, points.Count).All(i => Cross2(points[i], points[(i + 1) % points.Count], zero) >= -Eps)) return true;
            for (int i = 0; i < points.Count; i++)
            {
                var a = points[i]; var b = points[(i + 1) % points.Count];
                var dx = b.X - a.X; var dy = b.Y - a.Y; var ll = dx * dx + dy * dy;
                var t = ll <= Eps * Eps ? 0 : Math.Max(0, Math.Min(1, -(a.X * dx + a.Y * dy) / ll));
                if (Radial(new P2(a.X + t * dx, a.Y + t * dy), zero) <= cylinder.Radius + Eps) return true;
            }
            return false;
        }
        private static List<Point3Dto> ClipAxial(List<Point3Dto> input, Cylinder cylinder, double limit, double sign)
        {
            var output = new List<Point3Dto>();
            if (input.Count == 0) return output;
            var prev = input[input.Count - 1]; var pd = sign * (Dot(Sub(prev, cylinder.Origin), cylinder.Axis) - limit);
            foreach (var curr in input)
            {
                var cd = sign * (Dot(Sub(curr, cylinder.Origin), cylinder.Axis) - limit);
                if ((pd >= -Eps) != (cd >= -Eps)) output.Add(Add(prev, Sub(curr, prev), pd / (pd - cd)));
                if (cd >= -Eps) output.Add(curr);
                prev = curr; pd = cd;
            }
            return output;
        }
    }
}
