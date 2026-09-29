using System.Collections.Generic;
using Autodesk.Navisworks.Api;

namespace KPLN_NavisMcpBridge
{
    internal static partial class ClashService
    {
        private static bool IsCeilingPath(string path) =>
            HasPathCategory(path, "Потолки", "Ceilings");

        private static bool IsCeilingServicePath(string path) =>
            HasPathCategory(path,
                "Кабельные лотки", "Cable Trays", "Короба", "Conduits",
                "Трубы", "Pipes", "Воздуховоды", "Ducts",
                "Материалы изоляции труб", "Pipe Insulations", "Pipe Insulation",
                "Материалы изоляции воздуховодов", "Duct Insulations", "Duct Insulation",
                "Соединительные детали трубопроводов", "Pipe Fittings",
                "Соединительные детали воздуховодов", "Duct Fittings");

        private static SurfacePlaneDto DescribeCeilingPlane(ModelItem item,
            Dictionary<string, FootprintMesh> cache)
        {
            var mesh = ReadFootprintMesh(item, cache);
            var result = mesh.Error == null
                ? DuctCeilingFootprint.MeasurePlane(mesh.Points)
                : new SurfacePlaneDto { Reason = mesh.Error };
            result.MeshDiagnostics = mesh.Diagnostics;
            return result;
        }
    }
}
