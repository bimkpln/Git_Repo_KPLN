using System;
using System.Collections.Generic;
namespace Autodesk.Revit.DB {
 public enum StorageType { String,Integer,Double,ElementId }
 public class ElementId {
  public long Value; public ElementId(long value){Value=value;}
  public static readonly ElementId InvalidElementId=new ElementId(-1);
 }
 public class Parameter {
  public StorageType StorageType=StorageType.String;
  public bool HasValue=true,Fail;
  public string Text; public int Integer; public double Number; public ElementId Reference;
  private void Check(){if(Fail)throw new Exception("read failure");}
  public string AsString(){Check();return Text;}
  public int AsInteger(){Check();return Integer;}
  public double AsDouble(){Check();return Number;}
  public ElementId AsElementId(){Check();return Reference;}
 }
 public class Element {
  private static long next;
  public ElementId Id=new ElementId(++next),TypeId=ElementId.InvalidElementId;
  public Document Document;
  public string Name="Element";
  public List<Parameter> Parameters=new List<Parameter>();
  public Dictionary<string,List<Parameter>> Named=new Dictionary<string,List<Parameter>>();
  public List<ElementId> MaterialIds=new List<ElementId>(),PaintIds=new List<ElementId>();
  public bool FailMaterials;
  public ElementId GetTypeId(){return TypeId;}
  public IList<Parameter> GetParameters(string name){List<Parameter> p;return Named.TryGetValue(name,out p)?p:new List<Parameter>();}
  public ICollection<ElementId> GetMaterialIds(bool paint){if(FailMaterials)throw new Exception("materials failure");return paint?PaintIds:MaterialIds;}
  public void Set(string name,Parameter p){Named[name]=new List<Parameter>{p};}
 }
 public class ElementType:Element {}
 public class RevitLinkInstance:Element {}
 public class FamilyInstance:Element { public Element SuperComponent; }
 public class Material:Element {}
 public class PropertySetElement:Element {}
 public class AppearanceAssetElement:Element {}
 public class Document {
  public Dictionary<ElementId,Element> Elements=new Dictionary<ElementId,Element>();
  public Element GetElement(ElementId id){Element e;return Elements.TryGetValue(id,out e)?e:null;}
  public T Add<T>(T e) where T:Element {Elements[e.Id]=e;e.Document=this;return e;}
 }
 public class Face {public ElementId MaterialElementId;}
 public class Solid {public List<Face> Faces=new List<Face>();}
}
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 internal static class Entry {internal static string Id(Autodesk.Revit.DB.ElementId id){return id==null?"":id.Value.ToString();}}
}
namespace Autodesk.Revit.DB.Structure {
 public class Rebar:Autodesk.Revit.DB.Element {
  public Autodesk.Revit.DB.ElementId HostId=Autodesk.Revit.DB.ElementId.InvalidElementId;
  public bool FailHost;
  public Autodesk.Revit.DB.ElementId GetHostId(){if(FailHost)throw new System.Exception("host failure");return HostId;}
 }
 public class RebarInSystem:Autodesk.Revit.DB.Element {
  public Autodesk.Revit.DB.ElementId HostId=Autodesk.Revit.DB.ElementId.InvalidElementId;
  public Autodesk.Revit.DB.ElementId SystemId;
  public Autodesk.Revit.DB.ElementId GetHostId(){return HostId;}
 }
}
