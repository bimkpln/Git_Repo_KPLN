using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 internal sealed class HostCandidate {
  internal Element Element;
  internal MaterialInfo Materials;
 }
 internal sealed class Hosts:IDisposable {
  internal readonly Dictionary<ElementId,HostCandidate> Candidates=new Dictionary<ElementId,HostCandidate>();
  internal readonly HashSet<int> Classes=new HashSet<int>();
  internal readonly List<object> InventoryWarnings=new List<object>();
  private readonly Dictionary<ElementId,Shape> _geometry=new Dictionary<ElementId,Shape>();
  private readonly Materials _materials;
  private readonly Document _doc;
  internal int ConcreteCount;
  internal Hosts(Document doc) {
   _doc=doc; _materials=new Materials(doc);
   // Scan all loaded physical model categories, not a single view or a short category allowlist.
   foreach(var element in new FilteredElementCollector(doc).WhereElementIsNotElementType()) {
    if(Materials.IsMaterialResource(element)) continue;
    if(element.Category==null || element.Category.CategoryType!=CategoryType.Model || element.ViewSpecific ||
       element.Category.Id==new ElementId(BuiltInCategory.OST_Rebar) || element is RevitLinkInstance) continue;
    var info=_materials.ElementMaterials(element);
    foreach(var cls in info.Classes) Classes.Add(cls);
    if(info.Classes.Count>0) ConcreteCount++;
    if(info.Errors.Count>0) InventoryWarnings.Add(new{element_id=Entry.Id(element.Id),message=String.Join("; ",info.Errors)});
    if(info.Classes.Count>0 || info.Errors.Count>0)
     Candidates[element.Id]=new HostCandidate{Element=element,Materials=info};
   }
  }
  internal HostResult Find(Element rebar) {
   return DirectHost.Resolve(rebar, element=> {
    HostCandidate cached;
    return Candidates.TryGetValue(element.Id,out cached) ? cached.Materials : _materials.ElementMaterials(element);
   }, ()=>FindGeometry(rebar));
  }
  private HostResult FindGeometry(Element rebar) {
   var result=new HostResult(); bool uncertain=false;
   using(var shape=Geometry.Read(rebar)) {
    var box=rebar.get_BoundingBox(null);
    if(shape.Failed || shape.Solids.Count==0 || box==null) {
     result.Status="geometry_uncertain"; result.Notes.Add("Недоступна полная геометрия стержня.");
     result.Notes.AddRange(shape.Errors); return result;
    }
    using(var outline=new Outline(box.Min,box.Max))
    using(var collector=new FilteredElementCollector(_doc)) {
     var ids=collector.WhereElementIsNotElementType().WherePasses(new BoundingBoxIntersectsFilter(outline)).ToElementIds();
     foreach(var id in ids) {
      HostCandidate host;
      if(!Candidates.TryGetValue(id,out host) || id==rebar.Id) continue;
      Shape geometry;
      if(!_geometry.TryGetValue(id,out geometry)) _geometry[id]=geometry=Geometry.Read(host.Element);
      var notes=new List<string>(); var classes=new HashSet<int>(); var evidence=new List<Dictionary<string,object>>();
      double volume=0;
      if(geometry.Failed || geometry.Solids.Count==0) notes.Add("Не обработана геометрия основы "+Entry.Id(id));
      else {
       foreach(var solid in geometry.Solids) {
        var material=_materials.SolidMaterials(solid,host.Materials);
        if(material.Classes.Count==0 && material.Errors.Count==0) continue;
        try {
         double v=Geometry.Intersection(shape,solid);
         if(v<=Geometry.VolumeTolerance) continue;
         if(material.Errors.Count>0) notes.Add("Не прочитаны материалы пересекающегося тела "+Entry.Id(id)+": "+String.Join("; ",material.Errors));
         if(material.Classes.Count>0) {
          volume+=v; foreach(var cls in material.Classes) classes.Add(cls);
          evidence.AddRange(material.Evidence);
         }
        } catch(Exception ex) { notes.Add("Пересечение с основой "+Entry.Id(id)+": "+ex.Message); }
       }
      }
      if(notes.Count>0) { uncertain=true; result.Notes.AddRange(notes); }
      if(volume>Geometry.VolumeTolerance) {
       result.Candidates.Add(new Dictionary<string,object>{
        {"host_id",Entry.Id(id)},{"host_name",host.Element.Name},{"intersection_volume_ft3",volume},
        {"classes",classes.OrderBy(c=>c).ToArray()},{"materials",evidence}});
      }
     }
    }
   }
   if(uncertain) { result.Status="geometry_uncertain"; return result; }
   if(result.Candidates.Count==0) { result.Status="host_not_found"; return result; }
   var ordered=result.Candidates.OrderByDescending(c=>(double)c["intersection_volume_ft3"])
     .ThenBy(c=>long.Parse((string)c["host_id"])).ToArray();
   var selected=ordered[0]; var hostClasses=(int[])selected["classes"];
   result.Id=Candidates.Keys.First(id=>Entry.Id(id)==(string)selected["host_id"]);
   if(hostClasses.Length!=1) { result.Status="host_class_uncertain"; result.Notes.Add("В выбранной основе пересекаются материалы разных классов бетона."); }
   else { result.Class=hostClasses[0]; result.Status="ready"; }
   return result;
  }
  public void Dispose() { foreach(var shape in _geometry.Values) shape.Dispose(); }
 }
}