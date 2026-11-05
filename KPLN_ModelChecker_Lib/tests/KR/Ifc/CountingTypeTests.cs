using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
namespace Autodesk.Revit.DB {
 public enum StorageType { Integer,String,Double,ElementId }
 public class Parameter {
  public bool HasValue=true,Fail;
  public StorageType StorageType=StorageType.Integer;
  public int Integer; public string Text; public double Number;
  public int AsInteger(){if(Fail) throw new Exception("Read error"); return Integer;}
  public double AsDouble(){if(Fail) throw new Exception("Read error"); return Number;}
  public string AsString(){if(Fail) throw new Exception("Read error"); return Text;}
 }
 public class ElementId { public static readonly ElementId InvalidElementId=new ElementId(); }
 public class Document {
  public Element Type;
  public Element GetElement(ElementId id) { return Type; }
 }
 public class Element {
  public Document Document=new Document();
  public ElementId TypeId=ElementId.InvalidElementId;
  public ElementId GetTypeId() { return TypeId; }
  public List<Parameter> Values=new List<Parameter>();
  public IList<Parameter> GetParameters(string name) {
   if(name!="Орг.ИзделиеТипПодсчета") throw new Exception("Wrong parameter");
   return Values;
  }
 }
}
class CountingTypeTests {
 static int count;
 static void Test(Parameter parameter,bool expected,bool warned=false,bool duplicate=false) {
  var e=new Element();if(parameter!=null)e.Values.Add(parameter);if(duplicate)e.Values.Add(parameter);
  string warning;bool result=CountingTypeFilter.Exclude(e,out warning);
  if(result!=expected || (warning!=null)!=warned) throw new Exception("Unexpected exclusion or warning");
  count++;
 }
 static void TypeTest(Parameter instance,Parameter type,bool expected,bool warned=false,bool duplicateInstance=false,bool duplicateType=false,bool missingType=false) {
  var e=new Element {TypeId=new ElementId()};
  if(instance!=null) e.Values.Add(instance);
  if(duplicateInstance) e.Values.Add(instance);
  if(!missingType) {
   e.Document.Type=new Element();
   if(type!=null) e.Document.Type.Values.Add(type);
   if(duplicateType) e.Document.Type.Values.Add(type);
  }
  string warning;
  if(CountingTypeFilter.Exclude(e,out warning)!=expected || (warning!=null)!=warned)
   throw new Exception("Unexpected type fallback result");
  count++;
 }
 static void Main(){
  Test(null,false);
  Test(new Parameter{HasValue=false,Integer=0},false);
  Test(new Parameter{HasValue=false,Integer=2},false);
  Test(new Parameter{Integer=0},true);
  Test(new Parameter{Integer=1},false);
  Test(new Parameter{Integer=2},true);
  Test(new Parameter{Integer=-1},true);
  foreach(var text in new[]{null,"","  ","1"," 1 ","1.0"})
   Test(new Parameter{StorageType=StorageType.String,Text=text},false);
  foreach(var text in new[]{"0","2","изделие"})
   Test(new Parameter{StorageType=StorageType.String,Text=text},true);
  Test(new Parameter{StorageType=StorageType.Double,Number=1.0},false);
  Test(new Parameter{StorageType=StorageType.Double,Number=2.0},true);
  Test(new Parameter{StorageType=StorageType.Double,Number=double.NaN},false,true);
  Test(new Parameter{StorageType=StorageType.ElementId},false,true);
  Test(new Parameter{Fail=true},false,true);
  Test(new Parameter{Integer=2},false,true,true);
  TypeTest(null,new Parameter{Integer=2},true);
  TypeTest(null,new Parameter{Integer=1},false);
  TypeTest(null,new Parameter{Integer=0},true);
  TypeTest(new Parameter{HasValue=false},new Parameter{Integer=2},true);
  TypeTest(new Parameter{StorageType=StorageType.String,Text=" "},new Parameter{Integer=2},true);
  TypeTest(new Parameter{Integer=1},new Parameter{Integer=2},false);
  TypeTest(new Parameter{Integer=2},new Parameter{Integer=1},true);
  TypeTest(new Parameter{Integer=0},new Parameter{Integer=1},true);
  TypeTest(null,null,false);
  TypeTest(null,new Parameter{HasValue=false},false);
  TypeTest(null,new Parameter{StorageType=StorageType.String,Text=""},false);
  TypeTest(new Parameter{HasValue=false},new Parameter{HasValue=false},false);
  TypeTest(new Parameter{Fail=true},new Parameter{Integer=2},false,true);
  TypeTest(new Parameter{Integer=2},new Parameter{Integer=2},false,true,true);
  TypeTest(null,new Parameter{Fail=true},false,true);
  TypeTest(null,new Parameter{Integer=2},false,true,false,true);
  TypeTest(null,null,false,true,false,false,true);
  TypeTest(new Parameter{Integer=1},null,false,false,false,false,true);
  Console.WriteLine("PASS: "+count+" counting-type exclusion cases, including instance/type fallback, precedence, unset vs explicit zero and read failures.");
 }
}