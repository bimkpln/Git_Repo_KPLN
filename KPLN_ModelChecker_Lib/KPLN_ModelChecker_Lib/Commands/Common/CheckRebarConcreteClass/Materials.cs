using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckRebarConcreteClass {
 internal sealed class MaterialInfo {
  internal readonly HashSet<int> Classes=new HashSet<int>();
  internal readonly List<Dictionary<string,object>> Evidence=new List<Dictionary<string,object>>();
  internal readonly List<string> Errors=new List<string>();
  internal void Merge(MaterialInfo other) {
   foreach(var c in other.Classes) Classes.Add(c);
   Evidence.AddRange(other.Evidence); Errors.AddRange(other.Errors);
  }
 }
 internal sealed class Materials {
  private readonly Document _doc;
  private readonly Dictionary<ElementId,MaterialInfo> _cache=new Dictionary<ElementId,MaterialInfo>();
  internal Materials(Document doc) { _doc=doc; }
  internal MaterialInfo Material(ElementId id) {
   MaterialInfo result;
   if(_cache.TryGetValue(id,out result)) return result;
   result=new MaterialInfo(); _cache[id]=result;
   try {
    var m=_doc.GetElement(id) as Material;
    if(m==null) { result.Errors.Add("Материал "+global::KPLN_ModelChecker_Lib.Commands.CheckRebarConcreteClass.Id(id)+" не найден"); return result; }
    var ps=m.GetParameters("Мтрл.КодМатериала");
    if(ps.Count==0) return result;
    if(ps.Count!=1) { result.Errors.Add("Неоднозначный код материала "+global::KPLN_ModelChecker_Lib.Commands.CheckRebarConcreteClass.Id(id)); return result; }
    var p=ps[0]; if(!p.HasValue) return result;
    string code;
    if(p.StorageType==StorageType.String) code=p.AsString();
    else if(p.StorageType==StorageType.Integer) code=p.AsInteger().ToString(CultureInfo.InvariantCulture);
    else { result.Errors.Add("Неподдерживаемый тип кода материала "+global::KPLN_ModelChecker_Lib.Commands.CheckRebarConcreteClass.Id(id)); return result; }
    var cls=Rules.ClassFromCode(code);
    if(cls.HasValue) {
     result.Classes.Add(cls.Value);
     result.Evidence.Add(new Dictionary<string,object>{{"material_id",global::KPLN_ModelChecker_Lib.Commands.CheckRebarConcreteClass.Id(id)},{"material_name",m.Name},{"material_code",code},{"concrete_class",cls.Value}});
    }
   } catch(Exception ex) { result.Errors.Add("Материал "+global::KPLN_ModelChecker_Lib.Commands.CheckRebarConcreteClass.Id(id)+": "+ex.Message); }
   return result;
  }
  // These are material definitions/assets, never placed concrete host elements.
  internal static bool IsMaterialResource(Element element) {
   return element is PropertySetElement || element is AppearanceAssetElement || element is Material;
  }
  internal MaterialInfo ElementMaterials(Element host) {
   if(IsMaterialResource(host)) return new MaterialInfo();
   var result=new MaterialInfo(); var ids=new HashSet<ElementId>();
   foreach(bool paint in new[]{false,true}) {
    try { foreach(var id in host.GetMaterialIds(paint)) ids.Add(id); }
    catch(Exception ex) { result.Errors.Add("Материалы геометрии: "+ex.Message); }
   }
   Action<Element> assigned=element=>{
    if(element==null) return;
    try {
     foreach(Parameter p in element.Parameters) {
      try {
       if(p.StorageType!=StorageType.ElementId || !p.HasValue) continue;
       var id=p.AsElementId();
       if(id!=null && id!=ElementId.InvalidElementId && _doc.GetElement(id) is Material) ids.Add(id);
      } catch(Exception ex) { result.Errors.Add("Назначение материала: "+ex.Message); }
     }
    } catch(Exception ex) { result.Errors.Add("Параметры материалов: "+ex.Message); }
   };
   assigned(host);
   try { assigned(_doc.GetElement(host.GetTypeId())); }
   catch(Exception ex) { result.Errors.Add("Материалы типа: "+ex.Message); }
   foreach(var id in ids) result.Merge(Material(id));
   return result;
  }
  internal MaterialInfo SolidMaterials(Solid solid,MaterialInfo fallback) {
   var result=new MaterialInfo(); var ids=new HashSet<ElementId>();
   try {
    foreach(Face face in solid.Faces)
     if(face.MaterialElementId!=ElementId.InvalidElementId) ids.Add(face.MaterialElementId);
    if(ids.Count==0) return fallback;
    foreach(var id in ids) result.Merge(Material(id));
   } catch(Exception ex) { result.Errors.Add("Материалы тела: "+ex.Message); }
   return result;
  }
 }
}