using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace KPLN_NavisMcpBridge
{
    internal static partial class ClashService
    {
        private static bool IsFramePath(string path)
        {
            var p = (path ?? "").ToLowerInvariant();
            if (p.Contains("профнаст") || p.Contains("deck")) return false;
            return p.Contains("нескаркас_балка") || p.Contains("/ каркас несущий /") ||
                   p.Contains("/ structural framing /");
        }

        private static AxisGeometryDto DescribeCachedFrameGeometry(
            ModelItem item, ComApi.InwOaPath originalPath, BoundDto bound, Dictionary<string, AxisGeometryDto> cache)
        {
            if (item == null || bound == null) return MeshAxisFitter.Unavailable("no-item-bounds");
            string key;
            try { key = GeometryInstancePath.Read(originalPath.ArrayData); }
            catch (Exception ex) { return MeshAxisFitter.Unavailable("invalid-instance-path: " + ex); }
            AxisGeometryDto result;
            if (cache.TryGetValue(key, out result)) return result;
            result = DescribeFrameGeometry(item, originalPath, bound);
            cache[key] = result;
            return result;
        }

        private static AxisGeometryDto DescribeFrameGeometry(ModelItem item, ComApi.InwOaPath originalPath, BoundDto bound)
        {
            const int maxPoints = 30000;
            var diagnostics = new MeshAxisDiagnostics { PointLimit = maxPoints, ReportedItemBounds = bound };
            Func<string, AxisGeometryDto> failure = reason =>
            {
                var result = MeshAxisFitter.Unavailable(reason);
                result.Diagnostics = diagnostics;
                return result;
            };
            try
            {
                // A clash on an assembly must not acquire the axis of its combined parts.
                var leaves = item.HasGeometry ? new[] { item } :
                    item.DescendantsAndSelf.Where(x => x.HasGeometry).Take(2).ToArray();
                if (leaves.Length != 1) return failure("not-single-geometry-item");
                var path = item.HasGeometry ? originalPath : ComApiBridge.ToInwOaPath(leaves[0]);
                if (path == null) return failure("no-com-path");
                diagnostics.InstancePath = GeometryInstancePath.Read(path.ArrayData);
                int skipped;
                var fragments = GetInstanceFragments(path, out skipped);
                diagnostics.MatchedFragments = fragments.Count;
                diagnostics.SkippedInstanceFragments = skipped;
                // Fragment boxes belong to the selected instance, unlike a shared node's geometry box.
                var target = MeshAxisFitter.UnionBounds(fragments.Select(f => DescribeBound(f.GetWorldBox())));
                diagnostics.SelectedFragmentBounds = target;
                if (target == null) return failure("invalid-fragment-world-bounds");
                var row = new List<Point3Dto>();
                var column = new List<Point3Dto>();
                foreach (var fragment in fragments)
                {
                    var matrix = ReadMatrix(fragment.GetLocalToWorldMatrix());
                    if (matrix == null || matrix.Any(x => double.IsNaN(x) || double.IsInfinity(x)))
                        return failure("invalid-world-transform");
                    var collector = new PrimitivePointCollector(matrix, maxPoints - row.Count);
                    fragment.GenerateSimplePrimitives(ComApi.nwEVertexProperty.eNONE, collector);
                    diagnostics.GeneratedTriangleCount += collector.RowVectorPoints.Count / 3;
                    if (collector.Truncated)
                        return failure(collector.LimitReached ? "mesh-point-limit" : "mesh-invalid-vertex");
                    row.AddRange(collector.RowVectorPoints);
                    column.AddRange(collector.ColumnVectorPoints);
                }
                diagnostics.RowMeshBounds = MeshAxisFitter.BoundsOf(row);
                diagnostics.ColumnMeshBounds = MeshAxisFitter.BoundsOf(column);
                if (row.Count < 6) return failure("no-triangle-mesh");
                var fit = MeshAxisFitter.FitWorld(row, column, target);
                fit.Diagnostics = diagnostics;
                return fit;
            }
            catch (Exception ex)
            {
                return failure("mesh-read-error: " + ex);
            }
        }
    }
}
