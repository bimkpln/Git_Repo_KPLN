using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace KPLN_NavisMcpBridge
{
    public sealed class DuctSectionDto
    {
        public string ServiceKind;
        public string NominalSize;
        public string NominalSizeSource;
        public double? NominalWidth;
        public double? NominalHeight;
        public double? NominalDiameter;
        public double? InsulationThickness;
        public string Error;
    }

    internal static partial class ClashService
    {
        private sealed class FootprintMesh
        {
            public List<Point3Dto> Points;
            public string Error;
            public List<FootprintFragmentDiagnostics> Diagnostics = new List<FootprintFragmentDiagnostics>();
        }
        private static readonly string[] DuctSizeNames = { "Размер воздуховода", "Duct Size" };
        private static readonly string[] InsulationThicknessNames =
        {
            "КП_И_Толщина стенки", "Толщина изоляции", "Insulation Thickness"
        };
        private static bool HasPathCategory(string path, params string[] names) =>
            (path ?? "").Split(new[] { " / " }, StringSplitOptions.None).Any(s => names.Any(n => string.Equals(s.Trim(), n, StringComparison.OrdinalIgnoreCase)));

        private static ClashFootprintDto DescribeDuctCeilingFootprint(ModelItem first, ModelItem second,
            string path1, string path2, Point3Dto contact, Dictionary<string, FootprintMesh> cache)
        {
            var kind1 = DuctServiceKind(path1);
            var kind2 = DuctServiceKind(path2);
            var ductSide = kind1 != null && HasPathCategory(path2, "Потолки", "Ceilings") ? 1 :
                kind2 != null && HasPathCategory(path1, "Потолки", "Ceilings") ? 2 : 0;
            if (ductSide == 0) return null;
            var ductItem = ductSide == 1 ? first : second;
            var duct = ReadFootprintMesh(ductItem, cache);
            var section = ReadExactDuctSection(ductItem, ductSide == 1 ? kind1 : kind2);
            var ceiling = ReadFootprintMesh(ductSide == 1 ? second : first, cache);
            var result = duct.Error != null || ceiling.Error != null ?
                DuctCeilingFootprint.Failure("duct: " + duct.Error + "; ceiling: " + ceiling.Error) :
                DuctCeilingFootprint.Measure(duct.Points, ceiling.Points, contact);
            result.DuctSide = ductSide;
            result.ServiceKind = section.ServiceKind;
            result.NominalSize = section.NominalSize;
            result.NominalSizeSource = section.NominalSizeSource;
            result.NominalWidth = section.NominalWidth;
            result.NominalHeight = section.NominalHeight;
            result.NominalDiameter = section.NominalDiameter;
            result.InsulationThickness = section.InsulationThickness;
            result.DuctMeshDiagnostics = duct.Diagnostics;
            result.CeilingMeshDiagnostics = ceiling.Diagnostics;
            return result;
        }

        private static string DuctServiceKind(string path)
        {
            if (HasPathCategory(path, "Воздуховоды", "Ducts")) return "duct";
            if (HasPathCategory(path, "Материалы изоляции воздуховодов", "Duct Insulations", "Duct Insulation") ||
                HasPathCategory(path, "Изоляция воздуховода")) return "duct-insulation";
            return null;
        }

        private static DuctSectionDto ReadExactDuctSection(ModelItem item, string serviceKind)
        {
            if (serviceKind == null) return null;
            var result = new DuctSectionDto { ServiceKind = serviceKind };
            if (item == null) return result;
            try
            {
                foreach (var candidate in item.AncestorsAndSelf)
                {
                    foreach (var category in candidate.PropertyCategories)
                    {
                        if (!NameMatches(category.DisplayName, ObjectCategoryNames)) continue;
                        if (serviceKind == "duct")
                        {
                            var diameter = LengthFromCategory(category, DiameterNames);
                            var width = LengthFromCategory(category, WidthNames);
                            var height = LengthFromCategory(category, HeightNames);
                            if (diameter.HasValue && diameter.Value > 0)
                            {
                                result.NominalDiameter = diameter;
                                result.NominalSizeSource = "Object/Diameter";
                                return result;
                            }
                            if (width.HasValue && height.HasValue && width.Value > 0 && height.Value > 0)
                            {
                                result.NominalWidth = width;
                                result.NominalHeight = height;
                                result.NominalSizeSource = "Object/Height+Width";
                                return result;
                            }
                        }
                        else if (serviceKind == "duct-insulation")
                        {
                            if (!result.InsulationThickness.HasValue)
                                result.InsulationThickness = LengthFromCategory(category, InsulationThicknessNames);
                        }
                    }

                    if (serviceKind != "duct-insulation") continue;
                    var node = State.GetGUIPropertyNode(ComApiBridge.ToInwOaPath(candidate), true);
                    foreach (var obj in Enumerate(node.GUIAttributes()))
                    {
                        var attribute = obj as ComApi.InwGUIAttribute2;
                        if (attribute == null || !NameMatches(attribute.ClassUserName, ObjectCategoryNames)) continue;
                        var value = ValueFromGuiAttribute(attribute, DuctSizeNames);
                        if (string.IsNullOrWhiteSpace(value)) continue;
                        if (result.NominalSize == null)
                        {
                            result.NominalSize = value;
                            result.NominalSizeSource = "Object/Duct Size";
                        }
                    }
                    if (result.NominalSize != null && result.InsulationThickness.HasValue) return result;
                }
            }
            catch (Exception ex) { result.Error = ex.ToString(); }
            return result;
        }

        private static double? LengthFromCategory(PropertyCategory category, string[] names)
        {
            foreach (var property in category.Properties)
            {
                if (!NameMatches(property.DisplayName, names)) continue;
                try
                {
                    var value = property.Value;
                    if (value != null && value.IsDoubleLength) return value.ToDoubleLength();
                }
                catch { return null; }
            }
            return null;
        }

        private static FootprintMesh ReadFootprintMesh(ModelItem item, Dictionary<string, FootprintMesh> cache)
        {
            string key = null;
            var result = new FootprintMesh();
            try
            {
                if (item == null) throw new InvalidOperationException("missing-model-item");
                var itemPath = ComApiBridge.ToInwOaPath(item);
                key = GeometryInstancePath.Read(itemPath.ArrayData);
                if (cache.TryGetValue(key, out result)) return result;
                result = new FootprintMesh();
                var points = new List<Point3Dto>();
                var leaves = item.HasGeometry ? new[] { item } : item.DescendantsAndSelf.Where(x => x.HasGeometry).Take(2).ToArray();
                if (leaves.Length != 1) throw new InvalidOperationException("not-single-geometry-item");
                var path = ComApiBridge.ToInwOaPath(leaves[0]);
                int skipped;
                foreach (var fragment in GetInstanceFragments(path, out skipped))
                {
                    var matrix = ReadMatrix(fragment.GetLocalToWorldMatrix());
                    if (matrix == null) throw new InvalidOperationException("invalid-world-transform");
                    var collector = new PrimitivePointCollector(matrix, 30000 - points.Count);
                    fragment.GenerateSimplePrimitives(ComApi.nwEVertexProperty.eNONE, collector);
                    if (collector.Truncated) throw new InvalidOperationException("mesh-incomplete-or-point-limit");
                    var diagnostics = new FootprintFragmentDiagnostics
                    {
                        Transform = matrix,
                        WorldBounds = DescribeBound(fragment.GetWorldBox()),
                        RowMeshBounds = MeshAxisFitter.BoundsOf(collector.RowVectorPoints),
                        ColumnMeshBounds = MeshAxisFitter.BoundsOf(collector.ColumnVectorPoints),
                        TriangleCount = collector.RowVectorPoints.Count / 3
                    };
                    result.Diagnostics.Add(diagnostics);
                    string selection;
                    var selected = FootprintMeshValidation.Select(collector.RowVectorPoints,
                        collector.ColumnVectorPoints, matrix, diagnostics.WorldBounds, out selection);
                    diagnostics.Selection = selection;
                    if (selected == null) throw new InvalidOperationException(selection);
                    if (selected.Count == 0) continue;
                    points.AddRange(selected);
                }
                if (points.Count < 6) throw new InvalidOperationException("no-triangle-mesh");
                result.Points = points;
            }
            catch (Exception ex) { result.Error = ex.ToString(); }
            if (key != null) cache[key] = result;
            return result;
        }
    }
}
