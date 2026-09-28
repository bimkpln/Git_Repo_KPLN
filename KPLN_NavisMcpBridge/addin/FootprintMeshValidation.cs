using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_NavisMcpBridge
{
    public sealed class FootprintFragmentDiagnostics
    {
        public BoundDto WorldBounds;
        public BoundDto RowMeshBounds;
        public BoundDto ColumnMeshBounds;
        public double[] Transform;
        public string Selection;
        public int TriangleCount;
    }

    internal static class FootprintMeshValidation
    {
        internal static IList<Point3Dto> Select(IList<Point3Dto> row, IList<Point3Dto> column,
            double[] matrix, BoundDto worldBounds, out string selection)
        {
            if ((row == null || row.Count == 0) && (column == null || column.Count == 0))
            {
                selection = "no-triangle-mesh-skipped";
                return Array.Empty<Point3Dto>();
            }
            selection = "invalid-world-transform";
            if (matrix == null || matrix.Length != 16 || matrix.Any(x => !Finite(x)) ||
                Math.Abs(matrix[15] - 1) > 1e-12) return null;
            // A fragment world box may conservatively enclose a transformed local box.
            // It must contain the mesh, but need not share the mesh's extrema.
            bool rowOk = new[] { 3, 7, 11 }.All(i => Math.Abs(matrix[i]) <= 1e-12) && Contains(row, worldBounds);
            bool columnOk = new[] { 12, 13, 14 }.All(i => Math.Abs(matrix[i]) <= 1e-12) && Contains(column, worldBounds);
            if (!rowOk && !columnOk) { selection = "mesh-world-bounds-mismatch"; return null; }
            if (rowOk && columnOk && (row.Count != column.Count || row.Where((p, i) =>
                Math.Abs(p.X - column[i].X) + Math.Abs(p.Y - column[i].Y) + Math.Abs(p.Z - column[i].Z) > 1e-6).Any()))
            { selection = "ambiguous-world-transform"; return null; }
            selection = rowOk ? "affine-row-contained" : "affine-column-contained";
            return rowOk ? row : column;
        }

        private static bool Contains(IList<Point3Dto> points, BoundDto bound)
        {
            if (points == null || points.Count < 6 || points.Count % 3 != 0 || bound?.Min == null || bound.Max == null)
                return false;
            var lo = new[] { bound.Min.X, bound.Min.Y, bound.Min.Z };
            var hi = new[] { bound.Max.X, bound.Max.Y, bound.Max.Z };
            if (lo.Concat(hi).Any(x => !Finite(x)) || Enumerable.Range(0, 3).Any(i => hi[i] < lo[i])) return false;
            double tolerance = Math.Max(1e-5, Math.Sqrt(Enumerable.Range(0, 3).Sum(i => (hi[i] - lo[i]) * (hi[i] - lo[i]))) * 1e-5);
            return points.All(p => p != null && Finite(p.X) && Finite(p.Y) && Finite(p.Z) &&
                p.X >= lo[0] - tolerance && p.X <= hi[0] + tolerance &&
                p.Y >= lo[1] - tolerance && p.Y <= hi[1] + tolerance &&
                p.Z >= lo[2] - tolerance && p.Z <= hi[2] + tolerance);
        }

        private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
