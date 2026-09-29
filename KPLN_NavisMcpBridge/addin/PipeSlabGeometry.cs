using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_NavisMcpBridge
{
    public sealed class SlabLocalFacesDto
    {
        public string Source = "slab-local-faces-1";
        public bool Usable;
        public string Reason;
        public Point3Dto Normal;
        public double Thickness, NormalReliability;
        public double? NearestBroadFace, NearestEdgeFace;
        public int TriangleCount;
    }

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
        public SlabLocalFacesDto LocalFaces;
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
            const double angularTolerance = 1e-2;
            int orientationRejected = 0, hullRejected = 0, circleRejected = 0,
                lateralRejected = 0, radiiRejected = 0, lengthRejected = 0, areaRejected = 0;
            string topologyReason = null;
            var faces = UniqueFaces(mesh);
            var normals = new List<Point3Dto>();
            foreach (var f in faces)
                // Dense Navisworks tessellation may split one circular side into
                // hundreds of nearly coplanar facets. Coarse clustering is used
                // only to propose an axis; every original face is validated below.
                if (!normals.Any(n => Math.Abs(Dot(n, f.N)) > 1 - angularTolerance)) normals.Add(f.N);
            if (normals.Count < 5 || normals.Count > 256)
            { reason = "pipe-normal-count-outside-cylinder-range:" + normals.Count; return null; }
            var candidates = normals.ToList();
            var fittedNormalAxis = NormalPlaneAxis(faces);
            if (fittedNormalAxis != null &&
                !candidates.Any(n => Math.Abs(Dot(n, fittedNormalAxis)) > 1 - 1e-6))
                candidates.Add(fittedNormalAxis);
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
            var totalFaceArea = faces.Sum(f => f.Area);
            foreach (var axis in candidates)
            {
                var alignedArea = faces.Where(f => {
                    var d = Math.Abs(Dot(axis, f.N));
                    return d < angularTolerance || d > 1 - angularTolerance;
                }).Sum(f => f.Area);
                if (totalFaceArea <= Eps || alignedArea < totalFaceArea * .95)
                { orientationRejected++; continue; }
                var origin = mesh[0];
                var u = Unit(Cross(axis, Math.Abs(axis.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
                var v = Cross(axis, u);
                var hull = Hull(mesh.Select(p => Project(p, origin, u, v)).ToList());
                if (hull.Count < 8) { hullRejected++; continue; }
                // Circumcenter from well-separated hull vertices, then validate ALL
                // outer vertices and lateral faces. World coordinates are localized.
                var a = hull[0]; var b = hull[hull.Count / 3]; var c = hull[2 * hull.Count / 3];
                var det = 2 * (a.X * (b.Y - c.Y) + b.X * (c.Y - a.Y) + c.X * (a.Y - b.Y));
                if (Math.Abs(det) < Eps * Eps) { circleRejected++; continue; }
                var aa = a.X * a.X + a.Y * a.Y; var bb = b.X * b.X + b.Y * b.Y; var cc = c.X * c.X + c.Y * c.Y;
                var center = new P2((aa * (b.Y - c.Y) + bb * (c.Y - a.Y) + cc * (a.Y - b.Y)) / det,
                    (aa * (c.X - b.X) + bb * (a.X - c.X) + cc * (b.X - a.X)) / det);
                var radius = Math.Sqrt((a.X - center.X) * (a.X - center.X) + (a.Y - center.Y) * (a.Y - center.Y));
                var tolerance = Math.Max(1e-6, radius * .001);
                if (radius <= Eps || hull.Any(p => Math.Abs(Radial(p, center) - radius) > tolerance))
                { circleRejected++; continue; }
                var angles = hull.Select(p => Math.Atan2(p.Y - center.Y, p.X - center.X)).OrderBy(x => x).ToList();
                angles.Add(angles[0] + 2 * Math.PI);
                if (Enumerable.Range(0, angles.Count - 1).Any(i => angles[i + 1] - angles[i] > Math.PI / 4 + 1e-5))
                { circleRejected++; continue; }
                var lateral = faces.Where(f => Math.Abs(Dot(axis, f.N)) < angularTolerance).ToList();
                if (lateral.Count < 16) { lateralRejected++; continue; }
                var axisOrigin = Add(Add(origin, u, center.X), v, center.Y);
                var outerFaces = lateral.Where(f => new[] { f.A, f.B, f.C }.All(p =>
                    Math.Abs(Radial(Project(p, origin, u, v), center) - radius) <= tolerance)).ToList();
                if (outerFaces.Count < 16) { radiiRejected++; continue; }
                var axial = outerFaces.SelectMany(f => new[] { f.A, f.B, f.C })
                    .Select(p => Dot(Sub(p, axisOrigin), axis)).ToList();
                if (axial.Count == 0 || axial.Max() - axial.Min() <= Eps) { lengthRejected++; continue; }
                if (!CompleteOuterLateralSurface(outerFaces, axisOrigin, axis, axial.Min(), axial.Max(), out topologyReason))
                { areaRejected++; continue; }
                // Navisworks may tessellate one straight pipe with intermediate
                // axial rings. The topology check above requires all non-manifold
                // boundaries to be on the two real ends, so it accepts intermediate
                // rings while rejecting a missing face or a gap between pieces.
                if (found != null && Math.Abs(Dot(found.Axis, axis)) < 1 - 1e-6) return null;
                found = new Cylinder { Axis = axis, Origin = axisOrigin, Radius = radius, Low = axial.Min(), High = axial.Max() };
            }
            if (found == null) reason = "pipe-cylinder-fit-failed:" +
                "orientation=" + orientationRejected + ",hull=" + hullRejected +
                ",circle=" + circleRejected + ",lateral=" + lateralRejected +
                ",radii=" + radiiRejected + ",length=" + lengthRejected +
                ",area=" + areaRejected + ",topology=" + topologyReason;
            return found;
        }

        private sealed class EdgeUse
        {
            public Point3Dto A, B;
            public int Count;
        }

        private sealed class AxialBoundary
        {
            public EdgeUse Edge;
            public double Axial;
        }

        private static bool CompleteOuterLateralSurface(List<Face> faces, Point3Dto origin,
            Point3Dto axis, double low, double high, out string reason)
        {
            reason = null;
            // Navisworks may emit the same tessellation vertex with small
            // per-triangle coordinate noise. Retry only the topological weld,
            // up to 0.305 mm; engineering dimensions still use raw vertices.
            foreach (var weldTolerance in new[] { 1e-6, 1e-5, 1e-4, 1e-3 })
                if (CompleteOuterLateralSurface(faces, origin, axis, low, high,
                    weldTolerance, out reason)) return true;
            return false;
        }

        private static bool CompleteOuterLateralSurface(List<Face> faces, Point3Dto origin,
            Point3Dto axis, double low, double high, double weldTolerance, out string reason)
        {
            reason = null;
            var edges = new Dictionary<string, EdgeUse>();
            foreach (var face in faces)
            {
                var points = new[] { face.A, face.B, face.C };
                for (int i = 0; i < 3; i++)
                {
                    var a = points[i]; var b = points[(i + 1) % 3];
                    var key = string.Join(";", new[] { PointKey(a, weldTolerance), PointKey(b, weldTolerance) }
                        .OrderBy(x => x, StringComparer.Ordinal));
                    if (!edges.TryGetValue(key, out var use))
                    {
                        use = new EdgeUse { A = a, B = b };
                        edges[key] = use;
                    }
                    use.Count++;
                    if (use.Count > 2) { reason = "nonmanifold-edge"; return false; }
                }
            }
            // The fitted axis is derived from tessellated face normals, so the
            // same end ring can differ slightly in axial projection.
            var tolerance = Math.Max(1e-5, (high - low) * 1e-4);
            var u = Unit(Cross(axis, Math.Abs(axis.Z) < .9 ? P(0, 0, 1) : P(1, 0, 0)));
            var v = Cross(axis, u);
            int lowEdges = 0, highEdges = 0;
            var internalRings = new List<AxialBoundary>();
            var longitudinalSeams = new List<AxialBoundary>();
            foreach (var edge in edges.Values.Where(x => x.Count == 1))
            {
                var a = Dot(Sub(edge.A, origin), axis);
                var b = Dot(Sub(edge.B, origin), axis);
                if (Math.Abs(a - low) <= tolerance && Math.Abs(b - low) <= tolerance) lowEdges++;
                else if (Math.Abs(a - high) <= tolerance && Math.Abs(b - high) <= tolerance) highEdges++;
                else if (Math.Abs(a - b) <= tolerance)
                    internalRings.Add(new AxialBoundary { Edge = edge, Axial = (a + b) / 2 });
                else
                {
                    var aa = Math.Atan2(Dot(Sub(edge.A, origin), v), Dot(Sub(edge.A, origin), u));
                    var bb = Math.Atan2(Dot(Sub(edge.B, origin), v), Dot(Sub(edge.B, origin), u));
                    var delta = bb - aa;
                    while (delta > Math.PI) delta -= 2 * Math.PI;
                    while (delta < -Math.PI) delta += 2 * Math.PI;
                    if (Math.Abs(delta) > 1e-3) { reason = "internal-boundary-diagonal"; return false; }
                    var angle = aa + delta / 2;
                    while (angle < 0) angle += 2 * Math.PI;
                    while (angle >= 2 * Math.PI) angle -= 2 * Math.PI;
                    longitudinalSeams.Add(new AxialBoundary { Edge = edge, Axial = angle });
                }
            }
            if (lowEdges < 8 || highEdges < 8)
            { reason = "end-boundaries:" + lowEdges + "/" + highEdges; return false; }
            if (internalRings.Count > 0)
            {
                var groups = new List<List<AxialBoundary>>();
                var clusterTolerance = Math.Max(tolerance, weldTolerance * 2);
                foreach (var boundary in internalRings.OrderBy(x => x.Axial))
                {
                    if (groups.Count == 0 || boundary.Axial - groups[groups.Count - 1].Last().Axial > clusterTolerance)
                        groups.Add(new List<AxialBoundary>());
                    groups[groups.Count - 1].Add(boundary);
                }
                foreach (var group in groups)
                {
                    // A legitimate internal split contributes two complete
                    // coincident boundary rings (one from each mesh fragment).
                    // One open end, a missing sector or a real axial gap has
                    // only one circumference and must remain unsupported.
                    double angularLength = 0;
                    var angles = new List<double>();
                    foreach (var boundary in group)
                    {
                        var aa = Math.Atan2(Dot(Sub(boundary.Edge.A, origin), v), Dot(Sub(boundary.Edge.A, origin), u));
                        var bb = Math.Atan2(Dot(Sub(boundary.Edge.B, origin), v), Dot(Sub(boundary.Edge.B, origin), u));
                        var delta = bb - aa;
                        while (delta > Math.PI) delta -= 2 * Math.PI;
                        while (delta < -Math.PI) delta += 2 * Math.PI;
                        angularLength += Math.Abs(delta);
                        angles.Add(aa); angles.Add(bb);
                    }
                    angles.Sort();
                    if (angles.Count > 0) angles.Add(angles[0] + 2 * Math.PI);
                    var maxGap = angles.Count > 1 ? Enumerable.Range(0, angles.Count - 1)
                        .Max(i => angles[i + 1] - angles[i]) : 2 * Math.PI;
                    if (angularLength < 3.5 * Math.PI || maxGap > Math.PI / 4 + 1e-3)
                    {
                        reason = "unpaired-internal-ring:edges=" + group.Count +
                            ",turns=" + (angularLength / (2 * Math.PI)).ToString("0.###") +
                            ",gap=" + maxGap.ToString("0.###");
                        return false;
                    }
                }
            }
            if (longitudinalSeams.Count > 0)
            {
                var groups = new List<List<AxialBoundary>>();
                foreach (var seam in longitudinalSeams.OrderBy(x => x.Axial))
                {
                    var group = groups.FirstOrDefault(g => {
                        var d = Math.Abs(g[0].Axial - seam.Axial);
                        return Math.Min(d, 2 * Math.PI - d) <= 1e-3;
                    });
                    if (group == null) { group = new List<AxialBoundary>(); groups.Add(group); }
                    group.Add(seam);
                }
                foreach (var group in groups)
                {
                    var intervals = group.Select(x => new[] {
                        Dot(Sub(x.Edge.A, origin), axis), Dot(Sub(x.Edge.B, origin), axis) }).ToList();
                    var min = intervals.Min(x => Math.Min(x[0], x[1]));
                    var max = intervals.Max(x => Math.Max(x[0], x[1]));
                    var length = intervals.Sum(x => Math.Abs(x[1] - x[0]));
                    if (min > low + tolerance || max < high - tolerance ||
                        length < 1.9 * (high - low))
                    {
                        reason = "unpaired-longitudinal-seam:edges=" + group.Count +
                            ",coverage=" + (length / (high - low)).ToString("0.###");
                        return false;
                    }
                }
            }
            return true;
        }

        private static Point3Dto NormalPlaneAxis(List<Face> faces)
        {
            var matrix = new double[3, 3];
            foreach (var f in faces)
            {
                var n = new[] { f.N.X, f.N.Y, f.N.Z };
                for (int i = 0; i < 3; i++)
                    for (int j = 0; j < 3; j++) matrix[i, j] += f.Area * n[i] * n[j];
            }
            var vectors = new double[,] { { 1, 0, 0 }, { 0, 1, 0 }, { 0, 0, 1 } };
            for (int iteration = 0; iteration < 40; iteration++)
            {
                int p = 0, q = 1;
                for (int i = 0; i < 3; i++)
                    for (int j = i + 1; j < 3; j++)
                        if (Math.Abs(matrix[i, j]) > Math.Abs(matrix[p, q])) { p = i; q = j; }
                if (Math.Abs(matrix[p, q]) <= 1e-14 * Math.Max(1,
                    Math.Abs(matrix[0, 0]) + Math.Abs(matrix[1, 1]) + Math.Abs(matrix[2, 2]))) break;
                var angle = .5 * Math.Atan2(2 * matrix[p, q], matrix[q, q] - matrix[p, p]);
                var cs = Math.Cos(angle); var sn = Math.Sin(angle);
                var app = matrix[p, p]; var aqq = matrix[q, q]; var apq = matrix[p, q];
                matrix[p, p] = cs * cs * app - 2 * cs * sn * apq + sn * sn * aqq;
                matrix[q, q] = sn * sn * app + 2 * cs * sn * apq + cs * cs * aqq;
                matrix[p, q] = matrix[q, p] = 0;
                for (int k = 0; k < 3; k++)
                {
                    if (k != p && k != q)
                    {
                        var kp = matrix[k, p]; var kq = matrix[k, q];
                        matrix[k, p] = matrix[p, k] = cs * kp - sn * kq;
                        matrix[k, q] = matrix[q, k] = sn * kp + cs * kq;
                    }
                    var vp = vectors[k, p]; var vq = vectors[k, q];
                    vectors[k, p] = cs * vp - sn * vq;
                    vectors[k, q] = sn * vp + cs * vq;
                }
            }
            var index = new[] { 0, 1, 2 }.OrderBy(i => matrix[i, i]).First();
            var axis = P(vectors[0, index], vectors[1, index], vectors[2, index]);
            var length = Math.Sqrt(Dot(axis, axis));
            return length > Eps ? P(axis.X / length, axis.Y / length, axis.Z / length) : null;
        }

        private static double Radial(P2 p, P2 c) => Math.Sqrt((p.X - c.X) * (p.X - c.X) + (p.Y - c.Y) * (p.Y - c.Y));
        private static string PointKey(Point3Dto p) => PointKey(p, 1e-6);
        private static string PointKey(Point3Dto p, double tolerance) =>
            Convert.ToInt64(Math.Round(p.X / tolerance)) + "," +
            Convert.ToInt64(Math.Round(p.Y / tolerance)) + "," +
            Convert.ToInt64(Math.Round(p.Z / tolerance));
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
