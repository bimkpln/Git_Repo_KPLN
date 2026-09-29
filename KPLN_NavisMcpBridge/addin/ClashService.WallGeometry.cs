using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace KPLN_NavisMcpBridge
{
    public sealed class WallOwnerCandidate
    {
        public string Path;
        public string Name;
        public string ElementId;
    }

    public sealed class WallGeometryDiagnostics
    {
        public string OwnerPath;
        public string ElementId;
        public int GeometryItems;
        public int MatchedFragments;
        public int SkippedInstanceFragments;
        public int TriangleCount;
        public int PointLimit = 30000;
        public BoundDto MeshBounds;
        public List<WallOwnerCandidate> Ancestors = new List<WallOwnerCandidate>();
        public List<FootprintFragmentDiagnostics> Fragments = new List<FootprintFragmentDiagnostics>();
    }

    internal static partial class ClashService
    {
        private static AxisGeometryDto DescribeWallServiceAxis(ModelItem item, Dictionary<string, FootprintMesh> cache)
        {
            var mesh = ReadFootprintMesh(item, cache);
            return mesh.Error == null ? MeshAxisFitter.Fit(mesh.Points, "mesh-surface-pca-row") :
                MeshAxisFitter.Unavailable(mesh.Error);
        }

        private static PlanGeometryDto WallUnavailable(string reason, WallGeometryDiagnostics diagnostics)
        {
            return new PlanGeometryDto { Source = "wall-instance-mesh-1", Usable = false,
                Reason = reason, Diagnostics = diagnostics };
        }

        private static string OwnRevitElementId(ModelItem item)
        {
            // The API's element-id category is instance data; never use a type name/id
            // or an inherited GUI property to select a container of several walls.
            var property = item.PropertyCategories.FindPropertyByName(
                PropertyCategoryNames.RevitElementId, DataPropertyNames.RevitElementIdValue);
            return property == null ? null : PropertyText(property)?.Trim();
        }

        private static ModelItem ResolveWallOwner(ModelItem item, WallGeometryDiagnostics diagnostics)
        {
            ModelItem owner = null;
            string elementId = null;
            foreach (var candidate in item.AncestorsAndSelf)
            {
                var id = OwnRevitElementId(candidate);
                var path = GeometryInstancePath.Read(ComApiBridge.ToInwOaPath(candidate).ArrayData);
                diagnostics.Ancestors.Add(new WallOwnerCandidate
                    { Path = path, Name = candidate.DisplayName, ElementId = id });
                if (owner != null && id != elementId) break;
                if (!string.IsNullOrWhiteSpace(id))
                {
                    owner = candidate;
                    elementId = id;
                    diagnostics.OwnerPath = path;
                    diagnostics.ElementId = id;
                }
            }
            return owner;
        }

        private static PlanGeometryDto DescribeCachedPlanGeometry(ModelItem item, BoundDto itemBound,
            Point3Dto clashPoint, Dictionary<string, WallMeshAnalysis> cache)
        {
            var diagnostics = new WallGeometryDiagnostics();
            try
            {
                if (item == null) return WallUnavailable("missing-model-item", diagnostics);
                var owner = ResolveWallOwner(item, diagnostics);
                if (owner == null) return WallUnavailable("wall-element-owner-not-found", diagnostics);
                WallMeshAnalysis analysis;
                if (!cache.TryGetValue(diagnostics.OwnerPath, out analysis))
                {
                    analysis = ReadCompleteWallMesh(owner, diagnostics);
                    cache[diagnostics.OwnerPath] = analysis;
                }
                return DescribeLocalWallFaces(analysis, clashPoint);
            }
            catch (Exception ex)
            {
                return WallUnavailable("wall-owner-read-error: " + ex.Message, diagnostics);
            }
        }

        private static WallMeshAnalysis ReadCompleteWallMesh(ModelItem owner, WallGeometryDiagnostics diagnostics)
        {
            try
            {
                var points = new List<Point3Dto>();
                foreach (var leaf in owner.DescendantsAndSelf)
                {
                    // A nested, different Revit element proves that this owner is too broad.
                    var nestedId = OwnRevitElementId(leaf);
                    if (!string.IsNullOrWhiteSpace(nestedId) && nestedId != diagnostics.ElementId)
                        throw new InvalidOperationException("wall-owner-contains-other-elements");
                    if (!leaf.HasGeometry) continue;
                    diagnostics.GeometryItems++;
                    var path = ComApiBridge.ToInwOaPath(leaf);
                    var leafKey = GeometryInstancePath.Read(path.ArrayData);
                    if (leafKey != diagnostics.OwnerPath &&
                        !leafKey.StartsWith(diagnostics.OwnerPath + "/", StringComparison.Ordinal))
                        throw new InvalidOperationException("wall-leaf-outside-owner-instance");
                    int skipped;
                    var fragments = GetInstanceFragments(path, out skipped);
                    diagnostics.MatchedFragments += fragments.Count;
                    diagnostics.SkippedInstanceFragments += skipped;
                    foreach (var fragment in fragments)
                    {
                        var matrix = ReadMatrix(fragment.GetLocalToWorldMatrix());
                        if (matrix == null) throw new InvalidOperationException("invalid-world-transform");
                        var collector = new PrimitivePointCollector(matrix, diagnostics.PointLimit - points.Count);
                        fragment.GenerateSimplePrimitives(ComApi.nwEVertexProperty.eNONE, collector);
                        if (collector.Truncated) throw new InvalidOperationException(
                            collector.LimitReached ? "wall-mesh-point-limit" : "wall-mesh-invalid-vertex");
                        var detail = new FootprintFragmentDiagnostics
                        {
                            Transform = matrix, WorldBounds = DescribeBound(fragment.GetWorldBox()),
                            RowMeshBounds = MeshAxisFitter.BoundsOf(collector.RowVectorPoints),
                            ColumnMeshBounds = MeshAxisFitter.BoundsOf(collector.ColumnVectorPoints),
                            TriangleCount = collector.RowVectorPoints.Count / 3
                        };
                        diagnostics.Fragments.Add(detail);
                        string selection;
                        var selected = FootprintMeshValidation.Select(collector.RowVectorPoints,
                            collector.ColumnVectorPoints, matrix, detail.WorldBounds, out selection);
                        detail.Selection = selection;
                        if (selected == null) throw new InvalidOperationException(selection);
                        points.AddRange(selected);
                    }
                }
                diagnostics.TriangleCount = points.Count / 3;
                diagnostics.MeshBounds = MeshAxisFitter.BoundsOf(points);
                return DescribeCompleteWallMesh(points, diagnostics);
            }
            catch (Exception ex)
            {
                return new WallMeshAnalysis { Plan = WallUnavailable("wall-mesh-read-error: " + ex.Message, diagnostics) };
            }
        }

        private sealed class WallNormalArea
        {
            public double X, Y, Area;
        }

        private static WallMeshAnalysis DescribeCompleteWallMesh(IList<Point3Dto> points,
            WallGeometryDiagnostics diagnostics)
        {
            var triangles = new List<Triangle3>();
            var normals = new List<WallNormalArea>();
            Func<string, WallMeshAnalysis> failure = reason => new WallMeshAnalysis
                { Plan = WallUnavailable(reason, diagnostics) };
            if (points == null || points.Count < 6 || points.Count % 3 != 0)
                return failure("wall-no-triangle-mesh");
            for (int i = 0; i < points.Count; i += 3)
            {
                var triangle = new Triangle3 { A = points[i], B = points[i + 1], C = points[i + 2] };
                triangles.Add(triangle);
                var normal = Cross(Subtract(triangle.B, triangle.A), Subtract(triangle.C, triangle.A));
                double length = Math.Sqrt(Dot(normal, normal));
                double horizontal = Math.Sqrt(normal.X * normal.X + normal.Y * normal.Y);
                if (length <= 1e-12 || horizontal / length < 0.9999) continue;
                double nx = normal.X / horizontal, ny = normal.Y / horizontal;
                var group = normals.FirstOrDefault(x => Math.Abs(x.X * nx + x.Y * ny) > 1 - 1e-6);
                if (group == null) { group = new WallNormalArea { X = nx, Y = ny }; normals.Add(group); }
                group.Area += length * 0.5;
            }
            var sorted = normals.OrderByDescending(x => x.Area).ToArray();
            if (sorted.Length == 0) return failure("wall-no-vertical-faces");
            var broad = sorted[0];
            double reliability = broad.Area / Math.Max(sorted.Length > 1 ? sorted[1].Area : 0, 1e-12);
            if (reliability < 2) return failure("wall-broad-direction-ambiguous");
            // Face normals avoid PCA rotation from asymmetric tessellation or openings.
            double ax = -broad.Y, ay = broad.X;
            var origin = points[0];
            var along = points.Select(p => (p.X - origin.X) * ax + (p.Y - origin.Y) * ay).ToArray();
            var across = points.Select(p => (p.X - origin.X) * broad.X + (p.Y - origin.Y) * broad.Y).ToArray();
            double lo = across.Min(), hi = across.Max();
            if (hi - lo <= 1e-6) return failure("wall-zero-thickness-surface");
            bool lower = false, upper = false;
            foreach (var triangle in triangles)
            {
                var normal = Cross(Subtract(triangle.B, triangle.A), Subtract(triangle.C, triangle.A));
                double length = Math.Sqrt(Dot(normal, normal));
                if (length <= 1e-12 || Math.Abs((normal.X * broad.X + normal.Y * broad.Y) / length) < 1 - 1e-6) continue;
                double projection = (triangle.A.X - origin.X) * broad.X + (triangle.A.Y - origin.Y) * broad.Y;
                lower |= Math.Abs(projection - lo) <= 1e-5;
                upper |= Math.Abs(projection - hi) <= 1e-5;
            }
            if (!lower || !upper) return failure("wall-opposite-broad-faces-missing");
            double centerAlong = (along.Min() + along.Max()) * 0.5, centerAcross = (lo + hi) * 0.5;
            var plan = new PlanGeometryDto
            {
                Source = "wall-instance-mesh-1", Usable = true, Diagnostics = diagnostics,
                AxisX = ax, AxisY = ay, CenterX = origin.X + ax * centerAlong + broad.X * centerAcross,
                CenterY = origin.Y + ay * centerAlong + broad.Y * centerAcross,
                HalfLength = (along.Max() - along.Min()) * 0.5, HalfThickness = (hi - lo) * 0.5,
                Reliability = reliability, PointCount = points.Count
            };
            return new WallMeshAnalysis { Plan = plan, Triangles = triangles };
        }
    }
}
