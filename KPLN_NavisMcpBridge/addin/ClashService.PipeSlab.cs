using System;
using System.Collections.Generic;
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
                var pipe = ReadFootprintMesh(pipeItem, cache);
                var slab = ReadFootprintMesh(slabItem, cache);
                var result = pipe.Error != null || slab.Error != null ? new PipeSlabDto {
                    Reason = "pipe: " + pipe.Error + "; slab: " + slab.Error
                } : DuctCeilingFootprint.MeasurePipeSlab(pipe.Points, slab.Points, contact);
                result.PipeSide = side;
                result.PipeId = GeometryInstancePath.Read(ComApiBridge.ToInwOaPath(pipeItem).ArrayData);
                result.SlabId = GeometryInstancePath.Read(ComApiBridge.ToInwOaPath(slabItem).ArrayData);
                return result;
            }
            catch (Exception ex) { return new PipeSlabDto { PipeSide = side, Reason = ex.ToString() }; }
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
