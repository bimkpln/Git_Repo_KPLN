using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckIfcRebarMarks {
 internal sealed class MarkSource {
  internal Element Element;
  internal Parameter Parameter;
  internal string Value;
  internal string RebarValue;
  private static Parameter Read(Element element) {
   var parameters = element.GetParameters("Мрк.МаркаКонструкции");
   return parameters.Count == 1 && parameters[0].StorageType == StorageType.String ? parameters[0] : null;
  }
  internal static MarkSource Resolve(Element rebar) {
   var own = Read(rebar);
   var result = new MarkSource { Element=rebar, Parameter=own, Value=own?.AsString(), RebarValue=own?.AsString() };
   if(!String.IsNullOrWhiteSpace(result.Value)) return result;
   var seen = new HashSet<ElementId> { rebar.Id };
   for(var parent=(rebar as FamilyInstance)?.SuperComponent; parent!=null && seen.Add(parent.Id);
       parent=(parent as FamilyInstance)?.SuperComponent) {
    var parameter=Read(parent);
    if(parameter==null) continue;
    var value=parameter.AsString();
    // Retain an empty parameter if no filled mark exists anywhere in the chain.
    if(result.Parameter==null || !String.IsNullOrWhiteSpace(value)) {
     result.Element=parent; result.Parameter=parameter; result.Value=value;
    }
    if(!String.IsNullOrWhiteSpace(value)) return result;
   }
   return result;
  }
 }
}