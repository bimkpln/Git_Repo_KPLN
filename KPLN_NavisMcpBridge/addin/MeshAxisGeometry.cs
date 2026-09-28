using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_NavisMcpBridge
{
    public sealed class MeshAxisDiagnostics
    {
        public string InstancePath;
        public int MatchedFragments;
        public int SkippedInstanceFragments;
        public int GeneratedTriangleCount;
        public int PointLimit;
        public BoundDto ReportedItemBounds;
        public BoundDto SelectedFragmentBounds;
        public BoundDto RowMeshBounds;
        public BoundDto ColumnMeshBounds;
    }

    public sealed class AxisGeometryDto
    {
        public Point3Dto Axis;
        public Point3Dto Center;
        public double Length;
        public double TransverseSize;
        public double Reliability;
        public int TriangleCount;
        public string Source;
        public bool Usable;
        public string Reason;
        public MeshAxisDiagnostics Diagnostics;
    }

    internal static class MeshAxisFitter
    {
        internal static BoundDto BoundsOf(IEnumerable<Point3Dto> points)
        {
            if (points == null) return null;
            var lo = new Point3Dto { X = double.PositiveInfinity, Y = double.PositiveInfinity, Z = double.PositiveInfinity };
            var hi = new Point3Dto { X = double.NegativeInfinity, Y = double.NegativeInfinity, Z = double.NegativeInfinity };
            bool any = false;
            foreach (var p in points)
            {
                if (p == null || !Finite(p.X) || !Finite(p.Y) || !Finite(p.Z)) return null;
                any = true;
                lo.X = Math.Min(lo.X, p.X); lo.Y = Math.Min(lo.Y, p.Y); lo.Z = Math.Min(lo.Z, p.Z);
                hi.X = Math.Max(hi.X, p.X); hi.Y = Math.Max(hi.Y, p.Y); hi.Z = Math.Max(hi.Z, p.Z);
            }
            return any ? new BoundDto { Min = lo, Max = hi } : null;
        }

        internal static BoundDto UnionBounds(IEnumerable<BoundDto> bounds)
        {
            var points = new List<Point3Dto>();
            foreach (var b in bounds)
            {
                if (b?.Min == null || b.Max == null || b.Min.X > b.Max.X || b.Min.Y > b.Max.Y || b.Min.Z > b.Max.Z)
                    return null;
                points.Add(b.Min); points.Add(b.Max);
            }
            return BoundsOf(points);
        }

        internal static AxisGeometryDto Unavailable(string reason)
        {
            return new AxisGeometryDto { Source = "mesh-surface-pca", Reason = reason };
        }

        internal static AxisGeometryDto FitWorld(IList<Point3Dto> row, IList<Point3Dto> column, BoundDto bound)
        {
            var rowMatches = MatchesBounds(row, bound);
            var columnMatches = MatchesBounds(column, bound);
            if (!rowMatches && !columnMatches) return Unavailable("mesh-world-bounds-mismatch");
            var result = rowMatches ? Fit(row, "mesh-surface-pca-row") : Fit(column, "mesh-surface-pca-column");
            if (rowMatches && columnMatches)
            {
                var other = Fit(column, "mesh-surface-pca-column");
                if (result.Axis != null && other.Axis != null)
                {
                    var dot = result.Axis.X * other.Axis.X + result.Axis.Y * other.Axis.Y + result.Axis.Z * other.Axis.Z;
                    if (Math.Abs(dot) < 1 - 1e-6) return Unavailable("ambiguous-world-transform");
                }
            }
            return result;
        }

        internal static bool MatchesBounds(IList<Point3Dto> points, BoundDto bound)
        {
            if (points == null || points.Count == 0 || bound?.Min == null || bound.Max == null) return false;
            var lo = new[] { bound.Min.X, bound.Min.Y, bound.Min.Z };
            var hi = new[] { bound.Max.X, bound.Max.Y, bound.Max.Z };
            if (lo.Concat(hi).Any(x => !Finite(x))) return false;
            var min = new[] { double.PositiveInfinity, double.PositiveInfinity, double.PositiveInfinity };
            var max = new[] { double.NegativeInfinity, double.NegativeInfinity, double.NegativeInfinity };
            foreach (var p in points)
            {
                if (p == null) return false;
                var values = new[] { p.X, p.Y, p.Z };
                for (int i = 0; i < 3; i++)
                {
                    if (!Finite(values[i])) return false;
                    min[i] = Math.Min(min[i], values[i]); max[i] = Math.Max(max[i], values[i]);
                }
            }
            double error = 0, diagonal = 0;
            for (int i = 0; i < 3; i++)
            {
                if (hi[i] < lo[i]) return false;
                error += Math.Abs(min[i] - lo[i]) + Math.Abs(max[i] - hi[i]);
                diagonal += (hi[i] - lo[i]) * (hi[i] - lo[i]);
            }
            return error <= Math.Max(1e-5, Math.Sqrt(diagonal) * 1e-5);
        }

        internal static AxisGeometryDto Fit(IList<Point3Dto> vertices, string source)
        {
            if (vertices == null || vertices.Count < 6 || vertices.Count % 3 != 0)
                return Unavailable("no-complete-triangle-mesh");
            if (vertices.Any(p => p == null || !Finite(p.X) || !Finite(p.Y) || !Finite(p.Z)))
                return Unavailable("non-finite-mesh");

            // Integrate surface moments over each triangle; vertex density must not bias the axis.
            var origin = vertices[0];
            var first = new double[3];
            var second = new double[3, 3];
            double areaSum = 0;
            int count = 0;
            for (int i = 0; i < vertices.Count; i += 3)
            {
                var a = Relative(vertices[i], origin);
                var b = Relative(vertices[i + 1], origin);
                var c = Relative(vertices[i + 2], origin);
                var ab = new[] { b[0] - a[0], b[1] - a[1], b[2] - a[2] };
                var ac = new[] { c[0] - a[0], c[1] - a[1], c[2] - a[2] };
                var cross = new[] { ab[1] * ac[2] - ab[2] * ac[1],
                    ab[2] * ac[0] - ab[0] * ac[2], ab[0] * ac[1] - ab[1] * ac[0] };
                var area = Math.Sqrt(cross.Sum(x => x * x)) * 0.5;
                if (!Finite(area)) return Unavailable("non-finite-mesh-area");
                if (area <= 1e-16) continue;
                var sum = new[] { a[0] + b[0] + c[0], a[1] + b[1] + c[1], a[2] + b[2] + c[2] };
                areaSum += area;
                count++;
                for (int j = 0; j < 3; j++)
                {
                    first[j] += area * sum[j] / 3;
                    for (int k = 0; k < 3; k++)
                        second[j, k] += area * (sum[j] * sum[k] + a[j] * a[k] +
                            b[j] * b[k] + c[j] * c[k]) / 12;
                }
            }
            if (areaSum <= 1e-16 || count < 2) return Unavailable("degenerate-mesh");
            for (int j = 0; j < 3; j++) first[j] /= areaSum;
            for (int j = 0; j < 3; j++)
                for (int k = 0; k < 3; k++)
                {
                    second[j, k] = second[j, k] / areaSum - first[j] * first[k];
                    if (!Finite(second[j, k])) return Unavailable("non-finite-covariance");
                }

            var vectors = new double[,] { { 1, 0, 0 }, { 0, 1, 0 }, { 0, 0, 1 } };
            // Jacobi rotations diagonalize a real symmetric 3x3 covariance matrix.
            for (int iteration = 0; iteration < 40; iteration++)
            {
                int p = 0, q = 1;
                for (int j = 0; j < 3; j++)
                    for (int k = j + 1; k < 3; k++)
                        if (Math.Abs(second[j, k]) > Math.Abs(second[p, q])) { p = j; q = k; }
                if (Math.Abs(second[p, q]) <= 1e-14 * Math.Max(1,
                    Math.Abs(second[0, 0]) + Math.Abs(second[1, 1]) + Math.Abs(second[2, 2]))) break;
                var angle = 0.5 * Math.Atan2(2 * second[p, q], second[q, q] - second[p, p]);
                var cs = Math.Cos(angle);
                var sn = Math.Sin(angle);
                var app = second[p, p]; var aqq = second[q, q]; var apq = second[p, q];
                second[p, p] = cs * cs * app - 2 * cs * sn * apq + sn * sn * aqq;
                second[q, q] = sn * sn * app + 2 * cs * sn * apq + cs * cs * aqq;
                second[p, q] = second[q, p] = 0;
                for (int k = 0; k < 3; k++)
                {
                    if (k != p && k != q)
                    {
                        var kp = second[k, p]; var kq = second[k, q];
                        second[k, p] = second[p, k] = cs * kp - sn * kq;
                        second[k, q] = second[q, k] = sn * kp + cs * kq;
                    }
                    var vp = vectors[k, p]; var vq = vectors[k, q];
                    vectors[k, p] = cs * vp - sn * vq;
                    vectors[k, q] = sn * vp + cs * vq;
                }
            }
            var order = new[] { 0, 1, 2 }.OrderByDescending(i => second[i, i]).ToArray();
            var major = second[order[0], order[0]];
            if (major <= 1e-12) return Unavailable("degenerate-covariance");
            var extents = new double[3];
            for (int j = 0; j < 3; j++)
            {
                var lo = double.PositiveInfinity; var hi = double.NegativeInfinity;
                foreach (var vertex in vertices)
                {
                    var v = Relative(vertex, origin);
                    double projection = 0;
                    for (int k = 0; k < 3; k++) projection += v[k] * vectors[k, order[j]];
                    lo = Math.Min(lo, projection); hi = Math.Max(hi, projection);
                }
                extents[j] = hi - lo;
            }
            var axis = new[] { vectors[0, order[0]], vectors[1, order[0]], vectors[2, order[0]] };
            var signIndex = Enumerable.Range(0, 3).OrderByDescending(i => Math.Abs(axis[i])).First();
            if (axis[signIndex] < 0) for (int i = 0; i < 3; i++) axis[i] = -axis[i];
            var ratio = major / Math.Max(second[order[1], order[1]], 1e-12);
            var transverse = Math.Max(extents[1], extents[2]);
            // Numerical shape confidence, not an engineering tolerance or clash verdict.
            var usable = ratio >= 4 && extents[0] >= 2 * transverse;
            return new AxisGeometryDto
            {
                Axis = new Point3Dto { X = axis[0], Y = axis[1], Z = axis[2] },
                Center = new Point3Dto { X = origin.X + first[0], Y = origin.Y + first[1], Z = origin.Z + first[2] },
                Length = extents[0], TransverseSize = transverse, Reliability = ratio,
                TriangleCount = count, Source = source, Usable = usable,
                Reason = usable ? null : "ambiguous-principal-axis"
            };
        }

        private static bool Finite(double x) { return !double.IsNaN(x) && !double.IsInfinity(x); }
        private static double[] Relative(Point3Dto p, Point3Dto origin)
        {
            return new[] { p.X - origin.X, p.Y - origin.Y, p.Z - origin.Z };
        }
    }
}
