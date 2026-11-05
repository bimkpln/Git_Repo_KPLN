using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
namespace Autodesk.Revit.DB {
 public enum StorageType { String,Integer }
 public class ElementId {}
 public class Parameter {
  public StorageType StorageType=StorageType.String;
  public string Value;
  public string AsString() { if(StorageType!=StorageType.String) throw new Exception("Wrong storage"); return Value; }
 }
 public class Element {
  public ElementId Id=new ElementId();
  public List<Parameter> Parameters=new List<Parameter>();
  public IList<Parameter> GetParameters(string name) {
   if(name!="Мрк.МаркаКонструкции") throw new Exception("Wrong parameter");
   return Parameters;
  }
 }
 public class FamilyInstance:Element { public Element SuperComponent; }
}
class MarkSourceTests {
 static int count;
 static FamilyInstance Make(string value,Element parent=null,bool parameter=true) {
  var e=new FamilyInstance { SuperComponent=parent };
  if(parameter) e.Parameters.Add(new Parameter { Value=value });
  return e;
 }
 static void Check(bool ok,string message) { if(!ok) throw new Exception(message); count++; }
 static void Main() {
  var parent=Make("Фп-1");
  var bar=Make("",parent);
  var source=MarkSource.Resolve(bar);
  Check(source.Element==parent && source.Value=="Фп-1" && source.RebarValue=="","Empty child uses parent, retains raw value");
  Check(Decision.Evaluate(1,false,source.Parameter!=null,source.Value,"Фп-1")=="match","No false empty mark error");
  Check(Decision.Evaluate(1,false,source.Parameter!=null,source.Value,"Фп-2")=="mark_mismatch","Parent mismatch still reported");
  bar.Parameters[0].Value="Ст-1"; source=MarkSource.Resolve(bar);
  Check(source.Element==bar && source.Value=="Ст-1","Filled child has priority");
  bar.Parameters[0].Value="  "; source=MarkSource.Resolve(bar);
  Check(source.Element==parent,"Whitespace child uses parent");
  var middle=Make("",parent); bar=Make(null,middle,false); source=MarkSource.Resolve(bar);
  Check(source.Element==parent && source.Value=="Фп-1","Multiple nesting and missing parameter");
  middle.Parameters[0].Value="Б-1"; source=MarkSource.Resolve(bar);
  Check(source.Element==middle && source.Value=="Б-1","Nearest filled ancestor wins");
  parent.Parameters[0].Value=""; middle.Parameters[0].Value=""; source=MarkSource.Resolve(bar);
  Check(source.Parameter!=null && Decision.Evaluate(1,false,true,source.Value,"Фп-1")=="rebar_mark_empty","All empty remains error");
  bar=Make(null,null,false); source=MarkSource.Resolve(bar);
  Check(source.Parameter==null,"No parameter remains missing");
  parent=Make("Фп-1"); bar=Make("",parent); bar.Parameters[0].StorageType=StorageType.Integer;
  Check(MarkSource.Resolve(bar).Element==parent,"Non-string child safely uses parent");
  bar.Parameters.Add(new Parameter{Value="Ст-1"});
  Check(MarkSource.Resolve(bar).Element==parent,"Duplicate child parameters not chosen arbitrarily");
  bar=Make("",null); middle=Make("",bar); bar.SuperComponent=middle;
  Check(MarkSource.Resolve(bar).Value=="","Cycle terminates");
  Console.WriteLine("PASS: "+count+" mark source cases (Revit doubles).");
 }
}