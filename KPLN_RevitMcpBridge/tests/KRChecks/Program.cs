using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using KPLN_RevitMcpBridge.Services;
using KPLN_ModelChecker_User.Checks.KR;

internal static class Program
{
    private static int _assertions;
    private static void Check(bool value, string reason) { _assertions++; if (!value) throw new Exception(reason); }
    private static void Error(string code, Action action) {
        try { action(); throw new Exception("Expected " + code); }
        catch (BridgeException ex) { Check(ex.Code == code, "Wrong error: " + ex.Code + " != " + code); }
    }
    private static Dictionary<string, object> Obj(object value) { return Json.Parse(Json.Serialize(value)); }
    private static Dictionary<string, object> Input(string id = "CheckIfcRebarMarks") {
        return new Dictionary<string, object> {
            {"check_id",id}, {"expected_revision",7},
            {"library_module_id",Assembly.GetExecutingAssembly().ManifestModule.ModuleVersionId.ToString("D")}
        };
    }
    private static Dictionary<string, object> Page(string resultId, string section = "issues", int offset = 0) {
        return new Dictionary<string, object> {
            {"result_id",resultId}, {"section",section}, {"offset",offset}, {"limit",200}
        };
    }
    public static int Main() {
        try {
            var service = new KRCheckService(); var app = new UIApplication();
            var catalog = Json.Parse(Json.Serialize(new {items=service.Catalog()}));
            Check(((IList)catalog["items"]).Count == 2,"catalog");
            KRCheckAccess.Allowed=false;
            Error("kr_access_denied",()=>service.Run(app,Input(),"doc",7));
            var denied = (IList)Obj(new {items=service.Catalog()})["items"];
            Check(denied.Cast<Dictionary<string,object>>().All(x => Equals(x["available"],false)),"catalog denied");
            KRCheckAccess.Allowed=true; KRCheckAccess.Throws=true;
            Error("kr_access_denied",()=>service.Run(app,Input(),"doc",7));
            KRCheckAccess.Throws=false;
            Error("stale_document",()=>service.Run(app,Input(),"doc",8));
            var changed=Input();changed["library_module_id"]="wrong";
            Error("kr_library_changed",()=>service.Run(app,changed,"doc",7));
            app.ActiveUIDocument.Document.IsFamilyDocument=true;
            Error("invalid_document",()=>service.Run(app,Input(),"doc",7));
            app.ActiveUIDocument.Document.IsFamilyDocument=false;
            Error("unknown_kr_check",()=>service.Run(app,Input("arbitrary"),"doc",7));
            Check(FakeReport.Calls==0,"preconditions must prevent algorithm");
            var summary=Obj(service.Run(app,Input(),"doc",7));
            Check(!summary.ContainsKey("issues") && Equals(summary["status"],"partial"),"summary preserves status");
            Check((int)((Dictionary<string,object>)summary["section_counts"])["issues"]==201,"all counts");
            var resultId=(string)summary["result_id"];
            var first=Obj(service.Page(Page(resultId),"doc",7));
            Check(((IList)first["items"]).Count==200 && (bool)first["has_more"],"first page");
            var last=Obj(service.Page(Page(resultId,offset:200),"doc",8));
            Check(((IList)last["items"]).Count==1 && !(bool)last["has_more"] && (bool)last["is_stale"],"last page/stale");
            Check((long)Convert.ToInt64(last["revision_at_run"])==7,"snapshot revision");
            foreach(var section in new[]{"skipped_elements","run_warnings"})
                Check(((IList)Obj(service.Page(Page(resultId,section),"doc",7))["items"]).Count==1,"section "+section);
            Check(FakeReport.Calls==1,"paging cannot rerun check");
            Error("wrong_document",()=>service.Page(Page(resultId),"other",7));
            Error("invalid_input",()=>service.Page(Page(resultId,"invalid"),"doc",7));
            var oversized=Page(resultId);oversized["limit"]=201;
            Error("invalid_input",()=>service.Page(oversized,"doc",7));
            KRCheckAccess.Allowed=false;
            Error("kr_access_denied",()=>service.Page(Page(resultId),"doc",7));
            KRCheckAccess.Allowed=true;
            var concrete=Obj(service.Run(app,Input("CheckRebarConcreteClass"),"doc",7));
            Check(Equals(concrete["check_id"],"CheckRebarConcreteClass"),"dispatch concrete");
            FakeReport.Throws=true;
            Error("kr_check_failed",()=>service.Run(app,Input(),"doc",7));
            FakeReport.Throws=false;FakeReport.Override="broken";
            Error("kr_invalid_report",()=>service.Run(app,Input(),"doc",7));
            FakeReport.Override="{\"check_id\":\"other\",\"status\":\"completed\"}";
            Error("kr_invalid_report",()=>service.Run(app,Input(),"doc",7));
            FakeReport.Override=null;
            for(int n=0;n<4;n++) service.Run(app,Input(),"doc",7);
            Error("kr_result_expired",()=>service.Page(Page(resultId),"doc",7));
            Error("kr_result_expired",()=>new KRCheckService().Page(Page((string)concrete["result_id"]),"doc",7));
            var large="{\"text\":\""+new string('x',9*1024*1024)+"\"}";
            Error("invalid_json",()=>Json.Parse(large));
            Check(((string)Json.Parse(large,64*1024*1024)["text"]).Length==9*1024*1024,"internal report extended limit");
            Console.WriteLine("KR checks: "+_assertions+" assertions passed.");return 0;
        } catch(Exception ex) { Console.Error.WriteLine(ex);return 1; }
    }
}
