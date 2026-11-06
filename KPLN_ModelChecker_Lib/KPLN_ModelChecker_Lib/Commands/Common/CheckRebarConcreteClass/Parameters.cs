using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckRebarConcreteClass {
 internal sealed class ParameterValue {
  internal Element Source;
  internal string Text;
  internal string Error;
  internal bool Set => !String.IsNullOrWhiteSpace(Text);
 }
 internal static class Parameters {
  private static ParameterValue Read(Element element,string name) {
   var result=new ParameterValue{Source=element};
   if(element==null) return result;
   try {
    var values=element.GetParameters(name);
    if(values.Count==0) return result;
    if(values.Count!=1) { result.Error="Несколько параметров "+name; return result; }
    var p=values[0];
    if(!p.HasValue) return result;
    if(p.StorageType==StorageType.String) result.Text=p.AsString();
    else if(p.StorageType==StorageType.Integer) result.Text=p.AsInteger().ToString(CultureInfo.InvariantCulture);
    else if(p.StorageType==StorageType.Double) result.Text=p.AsDouble().ToString("R",CultureInfo.InvariantCulture);
    else result.Error="Неподдерживаемый тип параметра "+name;
   } catch(Exception ex) { result.Error=ex.Message; }
   return result;
  }
  internal static ParameterValue InstanceOrType(Element element,string name) {
   var own=Read(element,name);
   if(own.Set || own.Error!=null) return own;
   try {
    var typeId=element.GetTypeId();
    if(typeId==null || typeId==ElementId.InvalidElementId) return own;
    var type=element.Document.GetElement(typeId);
    if(type==null) { own.Error="Не найден тип элемента"; return own; }
    var typed=Read(type,name);
    return typed.Set || typed.Error!=null ? typed : own;
   } catch(Exception ex) { own.Error=ex.Message; return own; }
  }
  internal static bool RunningMeters(Element element,out string error) {
   var value=InstanceOrType(element,"Рзм.ПогМетрыВкл");
   error=value.Error;
   return error==null && value.Set && (value.Text.Trim()=="1" || String.Equals(value.Text.Trim(),"true",StringComparison.OrdinalIgnoreCase));
  }
  internal static ParameterValue RebarClass(Element rebar) {
   var seen=new HashSet<ElementId>();
   for(var element=rebar;element!=null && seen.Add(element.Id);element=(element as FamilyInstance)?.SuperComponent) {
    var value=InstanceOrType(element,"Арм.КлассБетона");
    if(value.Set || value.Error!=null) return value;
   }
   return new ParameterValue{Source=rebar};
  }
 }
}