using System;
using System.Collections.Generic;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
using Autodesk.Revit.DB;
namespace Autodesk.Revit.DB {
 public enum StorageType { String,Integer,ElementId,Double }
 public class ElementId { public static readonly ElementId InvalidElementId=new ElementId(); }
 public class Parameter {
  public StorageType StorageType=StorageType.String;
  public string Text; public int Number; public ElementId Reference; public bool Fail;
  public string AsString() { if(Fail) throw new Exception("parameter read failed"); return Text; }
  public int AsInteger() { if(Fail) throw new Exception("parameter read failed"); return Number; }
  public ElementId AsElementId() { if(Fail) throw new Exception("parameter read failed"); return Reference; }
 }
 public class Element {
  public ElementId Id=new ElementId(), TypeId=ElementId.InvalidElementId;
  public List<Parameter> Parameters=new List<Parameter>();
  public List<ElementId> Materials=new List<ElementId>(), Paint=new List<ElementId>();
  public bool FailMaterials;
  public ICollection<ElementId> GetMaterialIds(bool paint) { if(FailMaterials) throw new Exception("material list failed"); return paint ? Paint : Materials; }
  public ElementId GetTypeId() { return TypeId; }
 }
 public class Material:Element {
  public bool FailCode;
  public List<Parameter> Codes=new List<Parameter>();
  public IList<Parameter> GetParameters(string name) {
   if(name!="Мтрл.КодМатериала") throw new Exception("Wrong parameter");
   if(FailCode) throw new Exception("code unavailable");
   return Codes;
  }
 }
 public class Document {
  public Dictionary<ElementId,Element> Elements=new Dictionary<ElementId,Element>();
  public Element GetElement(ElementId id) { Element e; return Elements.TryGetValue(id,out e) ? e : null; }
 }
}
class MaterialCodeTests {
 static int count;
 static void Check(bool ok,string message) { if(!ok) throw new Exception(message); count++; }
 static Material Add(Document doc,Element host,string code) {
  var material=new Material();
  material.Codes.Add(new Parameter{Text=code});
  doc.Elements[material.Id]=material;
  host.Materials.Add(material.Id);
  return material;
 }
 static MaterialCodeState State(Document d,Element h) { return MaterialCodeFilter.Read(d,h).State; }
 static void Main() {
  Check(MaterialCodeFilter.IsConcreteCode("201234567"),"Concrete code");
  Check(MaterialCodeFilter.IsConcreteCode("200000000"),"Zero suffix");
  foreach(var value in new[]{null,"","20123456","2012345670","211234567","20123456A"," 201234567","201234567 ","201234567\n","20１２３４５６７"})
   Check(!MaterialCodeFilter.IsConcreteCode(value),"Invalid code "+value);
  var doc=new Document();var host=new Element();
  // Unrelated concrete in the project must not affect this host.
  Add(doc,new Element(),"201234567");
  Check(State(doc,host)==MaterialCodeState.NoMatch,"No materials excludes; unrelated concrete ignored");
  var m=Add(doc,host,"301234567");
  Check(State(doc,host)==MaterialCodeState.NoMatch,"Nonconcrete");
  m.Codes.Clear();Check(State(doc,host)==MaterialCodeState.NoMatch,"Absent code");
  m.Codes.Add(new Parameter{Text=""});Check(State(doc,host)==MaterialCodeState.NoMatch,"Empty code");
  m.Codes[0].Text="201234567";Check(State(doc,host)==MaterialCodeState.Concrete,"Assigned concrete");
  m.Codes[0].Fail=true;Check(State(doc,host)==MaterialCodeState.Unknown,"Cannot read code");
  m.Codes[0].Fail=false;m.Codes.Add(new Parameter{Text="201234567"});
  Check(State(doc,host)==MaterialCodeState.Unknown,"Ambiguous code");
  m.Codes.Clear();m.Codes.Add(new Parameter{StorageType=StorageType.Double});
  Check(State(doc,host)==MaterialCodeState.Unknown,"Unsupported code storage");
  m.Codes[0]=new Parameter{StorageType=StorageType.Integer,Number=201234567};
  Check(State(doc,host)==MaterialCodeState.Concrete,"Integer code");
  host.FailMaterials=true;
  Check(State(doc,host)==MaterialCodeState.Unknown,"Failed material enumeration");
  host.FailMaterials=false;
  m.Codes[0]=new Parameter{Text="301234567"};
  var concrete=Add(doc,host,"201234567");
  Check(State(doc,host)==MaterialCodeState.Concrete,"Any assigned material");
  host.Materials.Remove(concrete.Id);host.Paint.Add(concrete.Id);
  Check(State(doc,host)==MaterialCodeState.Concrete,"Paint assigned to candidate");
  host.Paint.Clear();
  host.Parameters.Add(new Parameter{StorageType=StorageType.ElementId,Reference=concrete.Id});
  Check(State(doc,host)==MaterialCodeState.Concrete,"Instance material assignment without solid");
  host.Parameters.Clear();
  var type=new Element();doc.Elements[type.Id]=type;host.TypeId=type.Id;
  type.Parameters.Add(new Parameter{StorageType=StorageType.ElementId,Reference=concrete.Id});
  Check(State(doc,host)==MaterialCodeState.Concrete,"Own type material assignment");
  type.Parameters.Clear();
  host.Materials.Add(new ElementId());
  Check(State(doc,host)==MaterialCodeState.Unknown,"Unresolvable material");
  host.Materials.Add(concrete.Id);
  Check(State(doc,host)==MaterialCodeState.Concrete,"Concrete still preserves warning despite another unreadable material");
  Console.WriteLine("PASS: "+count+" material-code and candidate-only reading cases (Revit doubles).");
 }
}