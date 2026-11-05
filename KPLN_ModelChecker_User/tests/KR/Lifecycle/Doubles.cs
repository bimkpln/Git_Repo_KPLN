using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.Checks.KR;
using KPLN_ModelChecker_User.WPFItems;

namespace Autodesk.Revit.Attributes {
 public enum TransactionMode { ReadOnly }
 public class TransactionAttribute:Attribute { public TransactionAttribute(TransactionMode mode){} }
}
namespace Autodesk.Revit.DB {
 public enum BuiltInCategory { OST_Rebar, OST_Walls }
 public class ElementId { public long Value; public ElementId(long n){Value=n;} }
 public class Element { public ElementId Id; public BuiltInCategory CategoryId; public Element(long id=1){Id=new ElementId(id);} }
 public class ElementType:Element { public ElementType(long id):base(id){} }
 public class Material:Element {}
 public class PropertySetElement:Element {}
 public class AppearanceAssetElement:Element {}
 public class ElementSet {}
 public class Document {
  public bool IsFamilyDocument;
  public List<Element> Elements=new List<Element>();
  public Element ProjectInformation=new Element(1);
  public Element GetElement(ElementId id)=>Elements.FirstOrDefault(e=>e.Id.Value==id.Value);
 }
 public class FilteredElementCollector:IDisposable,IEnumerable<Element> {
  IEnumerable<Element> elements;
  public FilteredElementCollector(Document doc){elements=doc.Elements;}
  public FilteredElementCollector OfCategory(BuiltInCategory category){elements=elements.Where(e=>e.CategoryId==category);return this;}
  public FilteredElementCollector WhereElementIsNotElementType(){elements=elements.Where(e=>!(e is ElementType));return this;}
  public IEnumerator<Element> GetEnumerator()=>elements.GetEnumerator(); IEnumerator IEnumerable.GetEnumerator()=>GetEnumerator();public void Dispose(){}
 }
}
namespace Autodesk.Revit.UI {
 public enum Result { Succeeded,Failed,Cancelled }
 public interface IExternalCommand { Result Execute(ExternalCommandData data,ref string message,ElementSet elements); }
 public class ExternalCommandData { public UIApplication Application; }
 public class UIDocument { public Document Document; }
 public class UIApplication { public UIDocument ActiveUIDocument; public IntPtr MainWindowHandle; }
 public enum TaskDialogCommonButtons { Ok }
 public class TaskDialog {
  public static int Calls;public static string Last; public string MainInstruction;
  public TaskDialog(string title){} public void Show(){Calls++;}
  public static void Show(string title,string message,TaskDialogCommonButtons buttons=TaskDialogCommonButtons.Ok){Calls++;Last=message;}
 }
}
namespace KPLN_Library_Forms.UI.HtmlWindow {
 public enum MessageType { Error,Warning,Success }
 public static class HtmlOutput { public static int Calls;public static void Print(string message,MessageType type){Calls++;} }
}
namespace KPLN_Library_DBWorker {
 public class User {public int SubDepartmentId=3;}
 public static class SQLiteMainService {public static User CurrentDBUser=new User();}
}
namespace KPLN_Library_ConfigWorker.Core {public interface IJsonSerializable {}}
namespace KPLN_Library_ExtensibleStorage {
 public enum MessageStatus {Ok,Error}
 public class ResultMessage {public string Description="";public MessageStatus CurrentStatus;}
}
namespace KPLN_ModelChecker_Lib.Common {
 public class FakeBuilder {
  public Guid Guid=Guid.Empty;
  public KPLN_Library_ExtensibleStorage.ResultMessage GetResMessage_Element(Element e)=>new KPLN_Library_ExtensibleStorage.ResultMessage();
 }
 public class ExtensibleStorageEntity {
  public string CheckName="legacy";public FakeBuilder ESBuilderRun=new FakeBuilder(), ESBuilderUserText=new FakeBuilder(),ESBuildergMarker=new FakeBuilder();
 }
}
namespace KPLN_ModelChecker_Lib {
 public enum ErrorStatus { Error,Warning }
 public class CheckerException:Exception {public CheckerException(string s):base(s){}}
 public class CheckCommandMsg {public Element MsgElement;public string Message;}
 public class CheckerEntity {
  public string Header,Description,Info;public Element[] Elements=new Element[0];public ErrorStatus Status;
  public CheckerEntity(string h,string d,string i){Header=h;Description=d;Info=i;}
  public CheckerEntity(object e,string h,string d,string i):this(h,d,i){Elements=((IEnumerable<Element>)e).ToArray();}
  public CheckerEntity Set_Status(ErrorStatus s){Status=s;return this;}
 }
}
namespace KPLN_Loader {public static class Application {public static Queue<object> OnIdling_CommandQueue=new Queue<object>();}}
namespace KPLN_ModelChecker_User.ExecutableCommand {public class ExcCmdWPFEntity_SetTimeRunLog {public ExcCmdWPFEntity_SetTimeRunLog(object b,DateTime d){}}}
namespace KPLN_ModelChecker_User.WPFItems {
 public class WPFEntity {public WPFEntity(CheckerEntity c,ExtensibleStorageEntity e){}}
 public class WPFReportCreator {public WPFReportCreator(WPFEntity[] e,string n,string d,string m=null){}public void SetWPFEntityFiltration_ByErrorHeader(){}}
}
namespace KPLN_ModelChecker_User.Forms {
 public class CheckMainForm:Window {public static int Calls;public CheckMainForm(UIApplication a,string n,Type t,WPFReportCreator r,bool b,object run,object text,object marker){}public new void Show(){Calls++;}}
 internal class KRCheckReportWindow {
  internal static int Calls;internal static KRCheckReport Last;internal static string LastTitle;
  internal KRCheckReportWindow(UIApplication a,Document d,string title,KRCheckReport report){Last=report;LastTitle=title;}
  public void Show(){Calls++;}
 }
}
namespace KPLN_ModelChecker_User.Common {
 public class CheckListAnnotationsSettingsM:KPLN_Library_ConfigWorker.Core.IJsonSerializable {}
 public static class CheckListAnnotationsService {public static string ApplyHighlightOrSelection(CheckListAnnotationsSettingsM c,UIApplication a,AbstrCheck b)=>"";}
}
internal static class Evaluation {
 internal static IDictionary<string,object> Data;internal static int Calls;internal static bool Fail;internal static Element[] Candidates;internal static Document Document;
 internal static Dictionary<string,object> Run(Document doc,Element[] candidates){Calls++;Candidates=candidates;Document=doc;if(Fail)throw new InvalidOperationException("algorithm failed");return new Dictionary<string,object>(Data);}
}
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc {
 public static class Checker {public const string CheckId="CheckIfcRebarMarks";internal static Dictionary<string,object> Run(Document d,Element[] e)=>Evaluation.Run(d,e);}
}
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 public static class Entry {public const string CheckId="CheckRebarConcreteClass";internal static Dictionary<string,object> Run(Document d,Element[] e)=>Evaluation.Run(d,e);}
}
