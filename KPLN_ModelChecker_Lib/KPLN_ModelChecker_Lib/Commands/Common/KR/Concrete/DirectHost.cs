using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 internal static class DirectHost {
  // System rebar exposes its physical host directly; IFC FamilyInstance uses geometry.
  internal static HostResult Resolve(Element rebar,Func<Element,MaterialInfo> materials,Func<HostResult> geometry) {
   if(!(rebar is Rebar) && !(rebar is RebarInSystem)) return geometry();
   var notes=new List<string>();
   try {
    var id=rebar is Rebar ? ((Rebar)rebar).GetHostId() : ((RebarInSystem)rebar).GetHostId();
    var host=id==null || id==ElementId.InvalidElementId ? null : rebar.Document.GetElement(id);
    if(host==null || host is ElementType || host is RevitLinkInstance || Materials.IsMaterialResource(host) || host.Id==rebar.Id)
     notes.Add("GetHostId не вернул доступную основу; использован поиск по геометрии.");
    else {
     var info=materials(host);
     if(info.Errors.Count==0 && info.Classes.Count==1) {
      var result=new HostResult{Status="ready",Mode="host",Id=host.Id,Class=info.Classes.Single()};
      result.Candidates.Add(new Dictionary<string,object>{
       {"host_id",Entry.Id(host.Id)},{"host_name",host.Name},{"classes",info.Classes.ToArray()},
       {"materials",info.Evidence},{"host_selection","GetHostId"}});
      return result;
     }
     notes.Add("Основа GetHostId "+Entry.Id(host.Id)+": "+
      (info.Errors.Count>0 ? "не удалось полностью прочитать материалы: "+String.Join("; ",info.Errors) :
       info.Classes.Count==0 ? "нет материала с бетонным кодом 20XXXXXXX" : "несколько классов бетона")+
      "; использован поиск по геометрии.");
    }
   } catch(Exception ex) {
    notes.Add("Ошибка чтения GetHostId/материалов: "+ex.Message+"; использован поиск по геометрии.");
   }
   var fallback=geometry();
   fallback.Notes.InsertRange(0,notes);
   return fallback;
  }
 }
}
