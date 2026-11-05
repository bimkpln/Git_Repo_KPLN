using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc {
 internal sealed class Shape : IDisposable {
  internal readonly List<Solid> Solids = new List<Solid>();
  internal bool Failed;
  internal readonly List<string> Errors = new List<string>();
  public void Dispose() { foreach(var s in Solids) s.Dispose(); }
 }
 internal static class Geometry {
  internal const double VolumeTolerance = 1e-9; // ft^3; numerical noise, not a containment percentage
  internal static Shape Read(Element e) {
   var result = new Shape();
   try {
    using(var options = new Options { DetailLevel = ViewDetailLevel.Fine, IncludeNonVisibleObjects = false })
    using(var geometry = e.get_Geometry(options)) {
     if(geometry != null) Walk(geometry, Transform.Identity, result);
    }
   } catch(Exception ex) { result.Failed = true; result.Errors.Add("Чтение геометрии: "+ex.GetType().Name+": "+ex.Message); }
   return result;
  }
  private static void Walk(GeometryElement geometry, Transform transform, Shape result) {
   foreach(var obj in geometry) {
    var solid = obj as Solid;
    if(solid != null && solid.Volume > VolumeTolerance) {
     try { result.Solids.Add(SolidUtils.CreateTransformed(solid, transform)); }
     catch(Exception ex) { result.Failed = true; result.Errors.Add("Преобразование геометрии: "+ex.GetType().Name+": "+ex.Message); }
    }
    var instance = obj as GeometryInstance;
    if(instance != null) {
     using(var nested = instance.GetSymbolGeometry()) {
      if(nested != null) Walk(nested, transform.Multiply(instance.Transform), result);
     }
    }
   }
  }
  // 1 concrete, 0 known other material, -1 unknown. Unknown never proves a match.
  internal static int MaterialKind(Document doc, ElementId id) {
   var m = doc.GetElement(id) as Material;
   if(m == null) return -1;
   var property = doc.GetElement(m.StructuralAssetId) as PropertySetElement;
   if(property != null) {
    using(var asset = property.GetStructuralAsset()) {
     if(asset != null) {
      if(asset.StructuralAssetClass == StructuralAssetClass.Concrete) return 1;
      var assetKind=asset.StructuralAssetClass.ToString();
      if(assetKind=="Metal" || assetKind=="Wood") return 0;
     }
    }
   }
   var label = (m.MaterialClass ?? "").Trim().ToLowerInvariant();
   if(label == "concrete" || label == "бетон") return 1;
   if(Regex.IsMatch(m.Name ?? "", @"(^|[_\s])бетон([\s_]|$)", RegexOptions.IgnoreCase)) return 1;
   if(new[]{"metal","steel","wood","металл","сталь","дерево","утеплитель","изоляция"}.Contains(label)) return 0;
   return -1;
  }
  internal static int SolidKind(Document doc, Element host, Solid solid) {
   var ids = new HashSet<ElementId>();
   foreach(Face f in solid.Faces) if(f.MaterialElementId != ElementId.InvalidElementId) ids.Add(f.MaterialElementId);
   if(ids.Count == 0) foreach(var id in host.GetMaterialIds(false)) ids.Add(id);
   if(ids.Count == 0) return -1;
   var kinds = ids.Select(id => MaterialKind(doc,id)).Distinct().ToArray();
   if(kinds.Length == 1) return kinds[0];
   return kinds.Contains(1) || kinds.Contains(-1) ? -1 : 0;
  }
  internal static double Intersection(Shape rebar, Solid host) {
   double volume = 0;
   foreach(var s in rebar.Solids)
    using(var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(s,host,BooleanOperationsType.Intersect))
     if(intersection != null) volume += Math.Abs(intersection.Volume);
   return volume;
  }
 }
}