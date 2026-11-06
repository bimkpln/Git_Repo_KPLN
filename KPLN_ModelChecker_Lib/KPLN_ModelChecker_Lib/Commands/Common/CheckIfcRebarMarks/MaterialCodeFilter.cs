using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckIfcRebarMarks {
 internal enum MaterialCodeState { NoMatch, Concrete, Unknown }
 internal sealed class MaterialCodeResult {
  internal MaterialCodeState State;
  internal readonly List<string> Errors=new List<string>();
 }
 internal static class MaterialCodeFilter {
  internal static bool IsConcreteCode(string value) {
   return value!=null && Regex.IsMatch(value,@"\A20[0-9]{7}\z");
  }
  internal static MaterialCodeResult Read(Document doc,Element host) {
   var result=new MaterialCodeResult();
   var materials=new Dictionary<ElementId,Material>();
   Action<Exception> failure=ex=>result.Errors.Add(ex.GetType().Name+": "+ex.Message);
   // Materials of this candidate only, including painted faces.
   foreach(bool paint in new[]{false,true}) {
    try {
     foreach(var id in host.GetMaterialIds(paint)) {
      var material=doc.GetElement(id) as Material;
      if(material==null) result.Errors.Add("Материал элемента не найден.");
      else materials[id]=material;
     }
    } catch(Exception ex) { failure(ex); }
   }
   // Void-only families may have assigned materials but no material-bearing solids.
   // Read material references on this instance and its type as well.
   Action<Element> assigned=element=>{
    if(element==null) return;
    try {
     foreach(Parameter parameter in element.Parameters) {
      try {
       if(parameter.StorageType!=StorageType.ElementId) continue;
       var id=parameter.AsElementId();
       if(id==null || id==ElementId.InvalidElementId) continue;
       var material=doc.GetElement(id) as Material;
       if(material!=null) materials[id]=material;
      } catch(Exception ex) { failure(ex); }
     }
    } catch(Exception ex) { failure(ex); }
   };
   assigned(host);
   try { assigned(doc.GetElement(host.GetTypeId())); } catch(Exception ex) { failure(ex); }
   foreach(var material in materials.Values) {
    try {
     var parameters=material.GetParameters("Мтрл.КодМатериала");
     if(parameters.Count==0) continue; // Readable absence is not a reading failure.
     if(parameters.Count!=1) { result.Errors.Add("Неоднозначный Мтрл.КодМатериала."); continue; }
     var parameter=parameters[0];
     string code;
     if(parameter.StorageType==StorageType.String) code=parameter.AsString();
     else if(parameter.StorageType==StorageType.Integer) code=parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
     else { result.Errors.Add("Неподдерживаемый тип Мтрл.КодМатериала."); continue; }
     if(IsConcreteCode(code)) { result.State=MaterialCodeState.Concrete; return result; }
    } catch(Exception ex) { failure(ex); }
   }
   result.State=result.Errors.Count>0 ? MaterialCodeState.Unknown : MaterialCodeState.NoMatch;
   return result;
  }
 }
}