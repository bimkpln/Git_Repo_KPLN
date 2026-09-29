using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using ClashApi = Autodesk.Navisworks.Api.Clash;

namespace KPLN_NavisMcpBridge
{
    internal static partial class ClashService
    {
        private static PipeSlabDto DescribePipeSlab(ModelItem first, ModelItem second,
            string path1, string path2, Point3Dto contact, Dictionary<string, FootprintMesh> cache)
        {
            var side = HasPathCategory(path1, "Трубы", "Pipes") && HasPathCategory(path2, "Перекрытия", "Floors") ? 1 :
                HasPathCategory(path2, "Трубы", "Pipes") && HasPathCategory(path1, "Перекрытия", "Floors") ? 2 : 0;
            if (side == 0) return null;
            try
            {
                var pipeItem = side == 1 ? first : second;
                var slabItem = side == 1 ? second : first;
                var slab = ReadFootprintMesh(slabItem, cache);
                // Same evidence as wall review: slab orientation and distance
                // from the contact to broad and edge/reveal faces. Pipe shape
                // validation is not a prerequisite for measuring slab faces.
                var result = new PipeSlabDto {
                    Source = "slab-local-faces-1",
                    Reason = "use-local-faces-and-item-bounds",
                    LocalFaces = slab.Error != null
                        ? new SlabLocalFacesDto { Reason = slab.Error }
                        : DescribeLocalSlabFaces(slab.Points, contact)
                };
                result.PipeSide = side;
                result.PipeId = GeometryInstancePath.Read(ComApiBridge.ToInwOaPath(pipeItem).ArrayData);
                result.SlabId = GeometryInstancePath.Read(ComApiBridge.ToInwOaPath(slabItem).ArrayData);
                return result;
            }
            catch (Exception ex) { return new PipeSlabDto { PipeSide = side, Reason = ex.ToString() }; }
        }

        private static SlabLocalFacesDto DescribeLocalSlabFaces(IList<Point3Dto> mesh, Point3Dto contact)
        {
            var result = new SlabLocalFacesDto();
            if (mesh == null || mesh.Count < 3 || mesh.Count % 3 != 0 || contact == null)
            { result.Reason = "missing-slab-mesh-or-contact"; return result; }
            var normals = new List<Point3Dto>();
            var weights = new List<double>();
            for (int i = 0; i < mesh.Count; i += 3)
            {
                var n = Cross(Subtract(mesh[i + 1], mesh[i]), Subtract(mesh[i + 2], mesh[i]));
                var weight = Math.Sqrt(Dot(n, n));
                if (weight <= 1e-12) continue;
                n = new Point3Dto { X = n.X / weight, Y = n.Y / weight, Z = n.Z / weight };
                int group = normals.FindIndex(x => Math.Abs(Dot(x, n)) > 1 - 1e-6);
                if (group < 0) { normals.Add(n); weights.Add(weight); }
                else weights[group] += weight;
            }
            var order = Enumerable.Range(0, weights.Count).OrderByDescending(i => weights[i]).ToList();
            if (order.Count < 2 || weights[order[0]] < 2 * weights[order[1]])
            { result.Reason = "ambiguous-slab-normal"; return result; }
            var normal = normals[order[0]];
            result.Normal = normal;
            result.NormalReliability = weights[order[0]] / weights[order[1]];
            var projections = mesh.Select(p => Dot(Subtract(p, contact), normal)).ToList();
            result.Thickness = projections.Max() - projections.Min();
            double broad = double.PositiveInfinity, edge = double.PositiveInfinity;
            for (int i = 0; i < mesh.Count; i += 3)
            {
                var a = mesh[i]; var b = mesh[i + 1]; var c = mesh[i + 2];
                var n = Cross(Subtract(b, a), Subtract(c, a));
                var length = Math.Sqrt(Dot(n, n));
                if (length <= 1e-12) continue;
                var distance = Math.Sqrt(PointTriangleDistanceSquared(contact, a, b, c));
                // As with wall faces, compare normal and in-plane components.
                if (Math.Abs(Dot(n, normal)) / length >= Math.Sqrt(.5))
                    broad = Math.Min(broad, distance);
                else edge = Math.Min(edge, distance);
                result.TriangleCount++;
            }
            if (!double.IsInfinity(broad)) result.NearestBroadFace = broad;
            if (!double.IsInfinity(edge)) result.NearestEdgeFace = edge;
            result.Usable = result.Thickness > 0 && result.NearestBroadFace.HasValue && result.NearestEdgeFace.HasValue;
            if (!result.Usable) result.Reason = "missing-broad-or-edge-faces";
            return result;
        }

        private static Dictionary<string, string> ReadStableGroupIds(string testName)
        {
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);
            var ambiguous = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                FindManagedTest(testName, out var test);
                CollectStableGroupIds(test, ids, ambiguous);
                foreach (var name in ambiguous) ids.Remove(name);
            }
            catch { ids.Clear(); } // Missing identity cannot establish a bundle.
            return ids;
        }
        private static void CollectStableGroupIds(GroupItem parent, Dictionary<string, string> ids, HashSet<string> ambiguous)
        {
            foreach (SavedItem item in parent.Children)
            {
                if (item is ClashApi.ClashResult result)
                {
                    var owner = CommentOwner(result) as ClashApi.ClashResultGroup;
                    if (ids.ContainsKey(result.DisplayName) || (owner != null && owner.Guid == Guid.Empty)) ambiguous.Add(result.DisplayName);
                    else ids[result.DisplayName] = owner != null && owner.Guid != Guid.Empty ? owner.Guid.ToString("D") : null;
                }
                if (item is GroupItem group) CollectStableGroupIds(group, ids, ambiguous);
            }
        }
    }
}
