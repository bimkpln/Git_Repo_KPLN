using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.Checks.KR;
using KPLN_ModelChecker_User.ExternalCommands;
using KPLN_ModelChecker_User.Forms;
using KPLN_Library_Forms.UI.HtmlWindow;

public class LegacyCheck:AbstrCheck {
 public static bool Fail,Empty,Warnings;
 public LegacyCheck(){ESEntity=new ExtensibleStorageEntity();}
 public override Element[] GetElemsToCheck()=>Empty?new Element[0]:new[]{new Element(2)};
 private protected override void CheckRElems_SetElemErrorColl(object[] elems){if(Warnings)PrepareElemsErrorColl.Add(new CheckCommandMsg{MsgElement=(Element)elems[0],Message="prepare"});}
 private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elems){
  if(Fail)throw new InvalidOperationException("legacy failed");
  if(Warnings)RunElemsWarningColl.Add(new CheckCommandMsg{Message="run"});
  _checkerEntitiesCollHeap.Add(new CheckerEntity("error","", ""));
  _checkerEntitiesCollHeap.Add(new CheckerEntity("warning","", "").Set_Status(ErrorStatus.Warning));
  return CheckResultStatus.Succeeded;
 }
}
public class LegacyCommand:AbstrCommand {}
internal static class Program {
 static int n;static void Check(bool v,string why){n++;if(!v)throw new Exception(why);}
 static Dictionary<string,object> Payload(string id,string status="partial") {
  return new Dictionary<string,object>{
   {"check_id",id},{"status",status},{"model_title","Модель"},{"checked_count",2},{"matched_count",0},{"error_count",1},{"review_count",1},{"unchecked_count",1},
   {"notes",new[]{"Закрытых рабочих наборов: 2"}},{"excluded_counting_type_count",3},
   {"issues",new object[]{new {rebar_id="5000000001",severity="error",status="class_mismatch",actual_class="25",expected_class=30,class_source_kind="type",class_source_id="99",element_ids=new[]{"5000000001","200","99"}},new {rebar_id="2",severity="review",status="geometry_uncertain",geometry_notes=new[]{"нет тела"}}}},
   {"skipped_elements",new[]{new {element_id="2",message="read failed"}}},{"run_warnings",new[]{new {message="material failed"}}}
  };
 }
 static void RunCheck<T>(UIApplication app,string id) where T:AbstrKRCheck,new(){
  Evaluation.Fail=false;Evaluation.Calls=0;Evaluation.Data=Payload(id);
  var expected=JObject.FromObject(Evaluation.Data);var check=new T();check.Set_UIAppData(app,app.ActiveUIDocument.Document);
  var candidates=check.GetElemsToCheck();Check(candidates.Length==2,"collect all rebar instances only");
  Check(check.ExecuteCheck(candidates,false)==CheckResultStatus.Succeeded,"partial means execution completed");
  Check(Evaluation.Calls==1 && ReferenceEquals(candidates,Evaluation.Candidates),"one evaluation with standard selected candidates");
  Check(ReferenceEquals(Evaluation.Document,app.ActiveUIDocument.Document),"correct document");
  Check(JToken.DeepEquals(expected,check.Report.Data),"all fields and statuses retained");
  Check(JToken.DeepEquals(expected,JObject.Parse(check.Report.ToJson())),"JSON contract retained");
  Check(check.CheckerEntitiesColl.Length==4,"all sections mapped");
  Check(check.CheckerEntitiesColl.Count(e=>e.Status==ErrorStatus.Error)==1,"uncertainties are warnings, not confirmed errors");
  Check(check.CheckerEntitiesColl[0].Elements.Any(e=>e.Id.Value==5000000001),"64-bit ID retained");
  Check(check.CheckerEntitiesColl.All(e=>e.Elements.All(x=>!(x is ElementType))),"type sources not selectable");
  Check(!check.CheckerEntitiesColl[0].Info.Contains("class_source_id"),"type ID not exposed as source instance");
  check.ExecuteCheck(candidates,true);Check(check.CheckerEntitiesColl.Length==1,"standard onlyErrorType projection");
  Check(JToken.DeepEquals(expected,check.Report.Data),"filter does not truncate full report");
  check.ExecuteCheck(candidates,false);Check(check.CheckerEntitiesColl.Length==4,"repeat does not accumulate findings");
  Evaluation.Data=Payload(id,"empty");Evaluation.Data["issues"]=new object[0];Evaluation.Data["skipped_elements"]=new object[0];Evaluation.Data["run_warnings"]=new object[0];
  Check(check.ExecuteCheck(new Element[0],false)==CheckResultStatus.NoItemsToCheck,"empty status retained");
  Check(check.Report.Status=="empty" && check.CheckerEntitiesColl.Length==0,"no synthetic error for empty model");
  Evaluation.Data=Payload(id);check.ExecuteCheck(new Element[0],false);
  Check(check.Report.Status=="partial" && Evaluation.Candidates.Length==0,"empty candidates still evaluate completeness");
  Evaluation.Fail=true;bool threw=false;try{check.ExecuteCheck(candidates,false);}catch(InvalidOperationException){threw=true;}
  Check(threw && check.Report==null && check.CheckerEntitiesColl.Length==0,"exception propagates; previous result cleared");
  Check(HtmlOutput.Calls==0,"no extra HTML windows");Evaluation.Fail=false;
 }
 [STAThread] public static int Main(){try{
  var doc=new Document();doc.Elements.AddRange(new Element[]{new Element(5000000001),new Element(2),new ElementType(99),new Element(200){CategoryId=BuiltInCategory.OST_Walls}});
  var app=new UIApplication{ActiveUIDocument=new UIDocument{Document=doc}};
  RunCheck<CheckIfcRebarMarks>(app,"CheckIfcRebarMarks");RunCheck<CheckRebarConcreteClass>(app,"CheckRebarConcreteClass");
  foreach(var command in new IExternalCommand[]{new CommandCheckKRIfcMarks(),new CommandCheckKRConcreteClass()}){
   string id=command is CommandCheckKRIfcMarks?"CheckIfcRebarMarks":"CheckRebarConcreteClass";
   Evaluation.Data=Payload(id);Evaluation.Fail=false;Evaluation.Calls=0;
   string message="";int before=KRCheckReportWindow.Calls;
   Check(command is AbstrCommand,"standard command base");
   Check(command.Execute(new ExternalCommandData{Application=app},ref message,new ElementSet())==Result.Succeeded,"manual run succeeds");
   Check(Evaluation.Calls==1 && KRCheckReportWindow.Calls==before+1,"one engine run and one agreed window");
   var original=KRCheckReport.Parse(JsonConvert.SerializeObject(Evaluation.Data));var current=KRCheckReportWindow.Last;
   Check(original.Summary==current.Summary && original.Notes==current.Notes,"UI summary/notes unchanged");
   Check(JsonConvert.SerializeObject(original.Rows)==JsonConvert.SerializeObject(current.Rows),"UI rows unchanged");
   Check(Mouse.OverrideCursor==null,"cursor restored");
   Check(HtmlOutput.Calls==0 && CheckMainForm.Calls==0 && KPLN_Loader.Application.OnIdling_CommandQueue.Count==0,"no alternate UI or model writes");
   var input=JObject.FromObject(Evaluation.Data);var json=id=="CheckIfcRebarMarks"?KPLN_ModelChecker_User.Checks.KR.Ifc.Checker.RunJson(app):KPLN_ModelChecker_User.Checks.KR.Concrete.Entry.RunJson(app);
   Check(JToken.DeepEquals(input,JObject.Parse(json)),"bridge entry point unchanged");
   Check(KRCheckReportWindow.Calls==before+1,"bridge does not show UI");
   Evaluation.Fail=true;
   Check(command.Execute(new ExternalCommandData{Application=app},ref message,new ElementSet())==Result.Failed && TaskDialog.Last.Contains("algorithm failed"),"same failure dialog");
   Check(KRCheckReportWindow.Calls==before+1 && Mouse.OverrideCursor==null,"failure does not show stale report; cursor restored");
   Evaluation.Fail=false;KPLN_Library_DBWorker.SQLiteMainService.CurrentDBUser.SubDepartmentId=2;int calls=Evaluation.Calls;
   Check(command.Execute(new ExternalCommandData{Application=app},ref message,new ElementSet())==Result.Cancelled && Evaluation.Calls==calls,"department denies before evaluation");
   KPLN_Library_DBWorker.SQLiteMainService.CurrentDBUser.SubDepartmentId=3;app.ActiveUIDocument=null;
   Check(command.Execute(new ExternalCommandData{Application=app},ref message,new ElementSet())==Result.Cancelled,"no project unchanged");app.ActiveUIDocument=new UIDocument{Document=doc};
  }
  LegacyCheck.Empty=true;var legacy=new LegacyCheck();legacy.Set_UIAppData(app,doc);
  Check(legacy.ExecuteCheck(legacy.GetElemsToCheck(),false)==CheckResultStatus.NoItemsToCheck && legacy.CheckerEntitiesColl.Length==1,"legacy empty behavior unchanged");
  LegacyCheck.Empty=false;LegacyCheck.Fail=true;int html=HtmlOutput.Calls;
  Check(legacy.ExecuteCheck(legacy.GetElemsToCheck(),false)==CheckResultStatus.Failed && HtmlOutput.Calls==html+1,"legacy failure reporting unchanged");
  LegacyCheck.Fail=false;LegacyCheck.Warnings=true;html=HtmlOutput.Calls;legacy=new LegacyCheck();legacy.Set_UIAppData(app,doc);
  Check(legacy.ExecuteCheck(legacy.GetElemsToCheck(),true)==CheckResultStatus.Succeeded && legacy.CheckerEntitiesColl.Length==1 && HtmlOutput.Calls==html+2,"legacy preparation/warnings/filter unchanged");
  LegacyCheck.Warnings=false;var oldCommand=new LegacyCommand();int oldViews=CheckMainForm.Calls;
  Check(oldCommand.ExecuteByUIApp<LegacyCheck>(app)==CheckResultStatus.Succeeded && CheckMainForm.Calls==oldViews+1,"legacy command still uses its report");
  Console.WriteLine("PASS: "+n+" lifecycle, legacy behavior, UI and bridge compatibility assertions.");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
