using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 public static class Entry {
  public const string CheckId="CheckRebarConcreteClass";
  public const string Title="Проверка класса бетона арматуры";
  internal static string Id(ElementId id) {
   if(id==null) return "";
   var property=typeof(ElementId).GetProperty("Value")??typeof(ElementId).GetProperty("IntegerValue");
   return Convert.ToString(property.GetValue(id),System.Globalization.CultureInfo.InvariantCulture);
  }
  internal static Dictionary<string,object> Run(Document doc, Element[] candidates) {
   if(doc==null || doc.IsFamilyDocument) throw new InvalidOperationException("Откройте проект RVT.");
   var started=DateTime.UtcNow;
   var issues=new List<Dictionary<string,object>>();
   var skipped=new List<object>(); var warnings=new List<object>(); var notes=new List<string>{
    "Проверяются системные и IFC-стержни текущей загруженной модели с Рзм.ПогМетрыВкл=1. Связи не проверяются.",
    "Орг.ИзделиеТипПодсчета: явно заданное значение экземпляра имеет приоритет; затем тип. Заданное значение, отличное от 1, исключает стержень.",
    "Арм.КлассБетона читается у экземпляра, затем типа; при пустом значении — у ближайшего родителя (экземпляр/тип).",
    "Класс основы: третья и четвёртая цифры Мтрл.КодМатериала из 9 цифр с началом 20.",
    "Системная арматура: сначала GetHostId; при недоступной основе/классе — геометрия. IFC: геометрия. При пересечениях выбирается максимальный объём; при равенстве — меньший ID."
   };
   int closed=doc.IsWorkshared ? new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).Count(w=>!w.IsOpen):0;
   int filtered=0,excluded=0,checkedCount=0,matched=0,errors=0,review=0;
   var bars=new List<Element>();
   using(var hosts=new Hosts(doc)) {
    bool fast=Rules.FastMode(hosts.Classes,hosts.InventoryWarnings.Count==0,closed);
    warnings.AddRange(hosts.InventoryWarnings);
    foreach(var bar in candidates) {
     string flagError;
     if(!Parameters.RunningMeters(bar,out flagError)) {
      if(flagError!=null) skipped.Add(new{element_id=Id(bar.Id),message="Не прочитан Рзм.ПогМетрыВкл: "+flagError});
      else filtered++;
      continue;
     }
     string countingError;
     if(CountingTypeFilter.Exclude(bar,out countingError)) { excluded++; continue; }
     if(countingError!=null) warnings.Add(new{element_id=Id(bar.Id),message=countingError});
     bars.Add(bar);
    }
    foreach(var bar in bars) {
     Dictionary<string,object> row;
     try { row=Inspect(bar,hosts,fast); }
     catch(Exception ex) { row=new Dictionary<string,object>{{"status","check_failed"},{"rebar_id",Id(bar.Id)},
      {"element_ids",new[]{Id(bar.Id)}},{"description",ex.GetType().Name+": "+ex.Message}}; }
     checkedCount++;
     string status=(string)row["status"];
     if(status=="match") { matched++; continue; }
     bool error=new[]{"class_mismatch","rebar_class_empty","rebar_class_invalid"}.Contains(status);
     if(error) errors++; else review++;
     row["severity"]=error?"error":"review"; row["text"]=row["description"];
     issues.Add(row);
    }
    if(closed>0) notes.Add("Закрытых рабочих наборов: "+closed+". Их содержимое не проверено; быстрый режим отключён.");
    if(hosts.InventoryWarnings.Count>0) notes.Add("Не удалось полностью прочитать материалы у "+hosts.InventoryWarnings.Count+" элементов. Быстрый режим отключён; подробности в run_warnings.");
    notes.Add(fast ? "Использован единый класс бетона проекта: "+hosts.Classes.Single()+". Поиск основы по геометрии не выполнялся."
     : "Для системной арматуры сначала использован GetHostId, затем при необходимости геометрия; для IFC — геометрия.");
    if(excluded>0) notes.Add("Исключено по Орг.ИзделиеТипПодсчета: "+excluded);
    var result=new Dictionary<string,object>{
     {"check_id",CheckId},{"title",Title},{"started_utc",started.ToString("o")},{"finished_utc",DateTime.UtcNow.ToString("o")},
     {"status",closed>0 || skipped.Count>0 || warnings.Count>0 || review>0 ? "partial" : bars.Count==0 ? "empty":"completed"},
     {"model_title",doc.Title},{"scope","loaded_host_document"},{"mode",fast?"single_class":"host_or_geometry"},
     {"project_concrete_classes",hosts.Classes.OrderBy(c=>c).ToArray()},{"concrete_element_count",hosts.ConcreteCount},
     {"collected_count",bars.Count},{"checked_count",checkedCount},{"matched_count",matched},{"error_count",errors},{"review_count",review},
     {"unchecked_count",bars.Count-checkedCount+skipped.Count},{"excluded_counting_type_count",excluded},
     {"excluded_running_meters_count",filtered},{"issue_count",issues.Count},{"issues",issues},{"skipped_elements",skipped},
     {"run_warnings",warnings},{"notes",notes}
    };
    return result;
   }
  }
  private static Dictionary<string,object> Inspect(Element bar,Hosts hosts,bool fast) {
   var actual=Parameters.RebarClass(bar);
   var parsed=Rules.ClassValue(actual.Text);
   var host=fast ? new HostResult{Status="ready",Mode="single_class",Class=hosts.Classes.Single()}:hosts.Find(bar);
   string status=actual.Error!=null ? "parameter_read_failed" : host.Status!="ready" ? host.Status :
    Rules.Compare(parsed,host.Class,!actual.Set);
   var messages=new Dictionary<string,string>{
    {"match","Класс бетона совпадает."},{"class_mismatch","Арм.КлассБетона не совпадает с классом бетона основы."},
    {"rebar_class_empty","Арм.КлассБетона не заполнен у стержня, типа и родителей."},
    {"rebar_class_invalid","Арм.КлассБетона не является положительным числом."},
    {"host_not_found","Не найдена пересекающаяся бетонная основа с кодом материала 20XXXXXXX."},
    {"host_class_uncertain","Класс бетона выбранной основы неоднозначен."},
    {"geometry_uncertain","Не удалось достоверно определить основу по геометрии/материалам."},
    {"parameter_read_failed","Ошибка чтения Арм.КлассБетона: "+actual.Error}
   };
   bool sourceIsType=actual.Source is ElementType;
   var ids=new[]{Id(bar.Id),sourceIsType ? "" : Id(actual.Source?.Id),Id(host.Id)}.Where(id=>!String.IsNullOrEmpty(id)).Distinct().ToArray();
   return new Dictionary<string,object>{
    {"status",status},{"description",messages.ContainsKey(status)?messages[status]:status},
    {"element_ids",ids},{"rebar_id",Id(bar.Id)},{"rebar_name",bar.Name},
    {"parent_id",Id((bar as FamilyInstance)?.SuperComponent?.Id)},
    {"actual_class",actual.Text},{"expected_class",host.Class},{"class_source_id",Id(actual.Source?.Id)},
    {"class_source_kind",actual.Source==null ? "unknown" : sourceIsType ? "type" : "instance"},
    {"class_source_name",actual.Source?.Name ?? ""},
    {"selected_host_id",Id(host.Id)},{"candidates",host.Candidates},{"geometry_notes",host.Notes},
    {"mode",host.Mode}
   };
  }
 }
}