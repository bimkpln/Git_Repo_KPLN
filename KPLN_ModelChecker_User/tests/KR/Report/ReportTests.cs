using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_ModelChecker_User.Checks.KR;
using KPLN_ModelChecker_User.Forms;
using Newtonsoft.Json.Linq;
class ReportTests {
 static int count;
 static void Check(bool condition,string name){count++;if(!condition)throw new Exception(name);}
 static JObject Sample(){return JObject.Parse(@"{
  'model_title':'Тестовая модель КР', 'status':'partial', 'checked_count':127, 'matched_count':124,
  'error_count':1,'review_count':2,'unchecked_count':1,'excluded_counting_type_count':8,
  'project_concrete_classes':[25,30],'mode':'host_or_geometry',
  'notes':['Область: текущая загруженная модель без связей.','Закрыт один рабочий набор — его содержимое не проверено.'],
  'issues':[
   {'status':'class_mismatch','severity':'error','rebar_id':'5000000001','element_ids':['5000000001','999','200'],
    'actual_class':'25,0','expected_class':30,'class_source_kind':'type','class_source_id':'999','class_source_name':'Арматура A500',
    'selected_host_id':'200','mode':'host','description':'Класс бетона арматуры не совпадает с классом материала основы.',
    'candidates':[{'host_id':'200','classes':[30],'materials':[{'material_id':'300','code':'203012345','class':30}]}]},
   {'status':'geometry_uncertain','severity':'review','rebar_id':'201','element_ids':['201'],
    'description':'Геометрия недоступна.','geometry_notes':['Не выполнено пересечение с основой 200.']}],
  'skipped_elements':[{'element_id':'202','message':'Ошибка чтения Рзм.ПогМетрыВкл'}],
  'run_warnings':[{'element_id':'203','message':'Не прочитаны материалы основы'}]
 }");}
 [STAThread] static int Main(string[] args){try{
  foreach(int id in new[]{3,8}) Check(KRCheckAccess.IsAllowedDepartment(id),"Allowed department");
  foreach(int id in new[]{0,1,2,4,5,6,7,9}) Check(!KRCheckAccess.IsAllowedDepartment(id),"Denied department");
  SQLiteMainService.User=null;Check(!KRCheckAccess.IsAllowed(),"Missing user");
  SQLiteMainService.Fail=true;Check(!KRCheckAccess.IsAllowed(),"DB failure");SQLiteMainService.Fail=false;
  SQLiteMainService.User=new User{SubDepartmentId=3};Check(KRCheckAccess.IsAllowed(),"KR user");
  var data=Sample();var report=KRCheckReport.Parse(data.ToString());
  Check(report.Rows.Count==4,"All result sections visible");
  Check(report.Summary.Contains("частично"),"Partial status");
  Check(report.Rows[0].SourceId=="" && !report.Rows[0].ElementIds.Contains("999"),"No type selection");
  Check(!report.Rows[0].Details.Contains("999") && report.Rows[0].Details.Contains("Арматура A500"),"Type name only in UI");
  Check(report.Rows[0].Actual=="«25,0»" && report.Rows[0].Expected=="«30»","Preserve values");
  Check(report.Rows[0].Element=="5000000001","Preserve 64 bit ID");
  Check(report.Rows[0].Details.Contains("203012345"),"Material evidence");
  Check(JToken.DeepEquals(data,JObject.Parse(report.Json)),"Full JSON preserved");
  var empty=Sample();empty["status"]="empty";Check(KRCheckReport.Parse(empty.ToString()).Summary.Contains("Нет элементов"),"Empty is not success");
  empty["status"]="completed";Check(KRCheckReport.Parse(empty.ToString()).Summary.Contains("Найдены ошибки"),"Completed with errors");
  empty["error_count"]=0;Check(KRCheckReport.Parse(empty.ToString()).Summary.Contains("Ошибок не найдено"),"Completed clean");
  var doc=new Document();doc.Elements[5000000001]=new Element();doc.Elements[999]=new ElementType();doc.Elements[300]=new Material();
  var ui=new UIDocument{Document=doc};var app=new UIApplication{ActiveUIDocument=ui};string status="";
  var handler=new KRReportSelectionHandler(doc,s=>status=s);
  handler.SetRequest(new[]{"5000000001","999","300","-1","bad","777"},true);handler.Execute(app);
  Check(ui.Selection.Ids.Count==1 && ui.Selection.Ids[0].Number==5000000001,"Select only existing instances / 64 bit");
  Check(ui.Shown.Count==1 && !handler.HasRequest,"Show and consume request");
  ui.Selection.Ids.Clear();ui.Document=new Document();handler.SetRequest(new[]{"5000000001"},false);handler.Execute(app);
  Check(ui.Selection.Ids.Count==0 && status.Contains("другая модель"),"Reject different document");
  ui.Document=doc;doc.IsValidObject=false;handler.SetRequest(new[]{"5000000001"},false);handler.Execute(app);
  Check(ui.Selection.Ids.Count==0,"Reject closed document");doc.IsValidObject=true;
  handler.IsClosed=true;handler.SetRequest(new[]{"5000000001"},false);handler.Execute(app);Check(ui.Selection.Ids.Count==0,"Ignore closed report");
  var window=new KRCheckReportWindow(app,doc,"КР: Проверка класса бетона арматуры",report);
  var visual=(FrameworkElement)window.Content;visual.Measure(new Size(1120,750));visual.Arrange(new Rect(0,0,1120,750));visual.UpdateLayout();
  Check(visual.ActualWidth>1000 && visual.ActualHeight>650,"Report layout");
  if(args.Length>0){var bitmap=new RenderTargetBitmap(1120,750,96,96,PixelFormats.Pbgra32);bitmap.Render(visual);
   var encoder=new PngBitmapEncoder();encoder.Frames.Add(BitmapFrame.Create(bitmap));using(var file=File.Create(args[0]))encoder.Save(file);}
  window.Close(); Console.WriteLine("PASS: "+count+" report, access, selection and layout assertions (Revit doubles).");return 0;
 }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
