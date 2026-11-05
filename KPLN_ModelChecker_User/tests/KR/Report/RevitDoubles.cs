using System;
using System.Collections.Generic;
namespace KPLN_Library_DBWorker {
 public class User { public int SubDepartmentId; }
 public static class SQLiteMainService {
  public static bool Fail; public static User User;
  public static User CurrentDBUser { get { if(Fail) throw new Exception("DB unavailable"); return User; } }
 }
}
namespace Autodesk.Revit.DB {
 public class ElementId { public long Number; public ElementId(long value){Number=value;} }
 public class Element { }
 public class ElementType:Element { }
 public class Material:Element { }
 public class PropertySetElement:Element { }
 public class AppearanceAssetElement:Element { }
 public class Document {
  public bool IsValidObject=true;
  public Dictionary<long,Element> Elements=new Dictionary<long,Element>();
  public Element GetElement(ElementId id){ return Elements.TryGetValue(id.Number,out var e)?e:null; }
 }
}
namespace Autodesk.Revit.UI {
 using Autodesk.Revit.DB;
 public interface IExternalEventHandler { void Execute(UIApplication app); string GetName(); }
 public enum ExternalEventRequest { Accepted, Pending, Denied }
 public class ExternalEvent:IDisposable {
  public bool IsPending;
  public static ExternalEvent Create(IExternalEventHandler handler){return new ExternalEvent();}
  public ExternalEventRequest Raise(){return ExternalEventRequest.Accepted;}
  public void Dispose() { }
 }
 public class UIApplication { public IntPtr MainWindowHandle; public UIDocument ActiveUIDocument; }
 public class UIDocument {
  public Document Document; public Selection Selection=new Selection(); public List<ElementId> Shown;
  public void ShowElements(List<ElementId> ids){Shown=ids;}
 }
 public class Selection { public List<ElementId> Ids=new List<ElementId>(); public void SetElementIds(List<ElementId> ids){Ids=ids;} }
}
