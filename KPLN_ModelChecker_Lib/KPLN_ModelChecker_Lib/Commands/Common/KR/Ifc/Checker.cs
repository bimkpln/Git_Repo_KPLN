using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc {
 public static class Checker {
  public const string CheckId = "CheckIfcRebarMarks";
  public const string Title = "Проверка марки IFC-арматуры";
  internal static string Id(ElementId id) {
   if(id == null) return "";
   var property = typeof(ElementId).GetProperty("Value") ?? typeof(ElementId).GetProperty("IntegerValue");
   return Convert.ToString(property.GetValue(id), System.Globalization.CultureInfo.InvariantCulture);
  }
  internal static Parameter UniqueParameter(Element e, string name) {
   if(e == null) return null;
   var ps = e.GetParameters(name);
   return ps.Count == 1 ? ps[0] : null;
  }
  internal static Dictionary<string,object> Run(Document doc, Element[] candidates) {
   if(doc == null || doc.IsFamilyDocument) throw new InvalidOperationException("Откройте проект RVT.");
   var started = DateTime.UtcNow;
   var issues = new List<Dictionary<string,object>>();
   var notes = new List<string> { "Область: вся загруженная текущая модель; связи не проверяются.",
    "Если Мрк.МаркаКонструкции стержня пустая или недоступна, используется ближайший родитель с заполненным текстовым параметром.",
    "Сравнение строк точное, включая пробелы и регистр. При нескольких бетонных основах выбирается наибольший объём пересечения; при точно равных объёмах — меньший ID.",
    "Бетон при успешной обработке определяется по физическому материалу, классу материала или имени «Бетон».",
    "Предупреждения о необработанном кандидате учитываются, если у его материала Мтрл.КодМатериала — ровно 9 цифр с началом 20 либо материалы/коды не удалось прочитать. При отсутствии подходящих кодов кандидат исключается." };
   var all = candidates;
   var bars = new List<Element>(); int unidentified = 0, excludedCountingType = 0, countingReadWarnings = 0;
   foreach(var e in all) {
    var flag = UniqueParameter(doc.GetElement(e.GetTypeId()), "Арм.ВыполненаСемейством");
    if(flag == null || flag.StorageType != StorageType.Integer) { unidentified++; continue; }
    if(flag.AsInteger() != 1) continue;
    string countingWarning;
    if(CountingTypeFilter.Exclude(e,out countingWarning)) { excludedCountingType++; continue; }
    if(countingWarning!=null) countingReadWarnings++;
    bars.Add(e);
   }
   if(excludedCountingType>0) notes.Add("Исключено IFC-стержней по явно заданному Орг.ИзделиеТипПодсчета, отличному от 1: "+excludedCountingType);
   if(countingReadWarnings>0) notes.Add("Не удалось однозначно прочитать Орг.ИзделиеТипПодсчета у "+countingReadWarnings+" стержней; они оставлены в проверке.");
   var categories = new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_StairsRuns, BuiltInCategory.OST_StairsLandings };
   var cache = new Dictionary<ElementId,Shape>();
   var materialCodes = new Dictionary<ElementId,MaterialCodeResult>();
   int processed = 0, matched = 0, errors = 0, review = 0;
   try {
    foreach(var bar in bars) {
     Dictionary<string,object> row;
     try { row = Inspect(doc, bar, categories, cache, materialCodes); }
     catch(Exception ex) { row = new Dictionary<string,object> { ["element_ids"]=new[]{Id(bar.Id)}, ["rebar_id"]=Id(bar.Id),
       ["status"]="check_failed", ["description"]="Ошибка обработки элемента: " + ex.GetType().Name + ": " + ex.Message }; }
     processed++;
     var status = (string)row["status"];
     if(status == "match") matched++;
     else {
      if(new[]{"mark_mismatch","rebar_mark_empty","host_mark_empty","parameter_missing"}.Contains(status)) errors++;
      else review++;
      row["text"] = status; row["severity"] = new[]{"mark_mismatch","rebar_mark_empty","host_mark_empty","parameter_missing"}.Contains(status) ? "error" : "review";
      issues.Add(row);
     }
    }
   } finally { foreach(var shape in cache.Values) shape.Dispose(); }
   int closed = doc.IsWorkshared ? new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).Count(w=>!w.IsOpen) : 0;
   if(closed > 0) notes.Add("Закрытых рабочих наборов: " + closed + ". Их содержимое не проверено.");
   if(unidentified > 0) notes.Add("Не определён признак IFC у " + unidentified + " элементов категории арматуры (нет однозначного параметра типа).");
   var result = new Dictionary<string,object> {
    ["check_id"]=CheckId, ["title"]=Title, ["started_utc"]=started.ToString("o"), ["finished_utc"]=DateTime.UtcNow.ToString("o"),
    ["status"]=(review>0 || closed>0 || unidentified>0 || processed<bars.Count) ? "partial" : bars.Count==0 ? "empty" : "completed",
    ["model_title"]=doc.Title, ["scope"]="loaded_host_document", ["collected_count"]=bars.Count, ["checked_count"]=processed,
    ["excluded_counting_type_count"]=excludedCountingType, ["ifc_collected_count"]=bars.Count+excludedCountingType,
    ["matched_count"]=matched, ["error_count"]=errors, ["review_count"]=review, ["unchecked_count"]=bars.Count-processed,
    ["issue_count"]=issues.Count, ["issues"]=issues, ["skipped_elements"]=new object[0], ["run_warnings"]=new object[0], ["notes"]=notes
   };
   return result;
  }
  private static Dictionary<string,object> Inspect(Document doc, Element bar, List<BuiltInCategory> cats, Dictionary<ElementId,Shape> cache, Dictionary<ElementId,MaterialCodeResult> materialCodes) {
   var parent = (bar as FamilyInstance)?.SuperComponent;
   var ancestors = new List<string>(); var seen = new HashSet<ElementId>();
   for(var p=parent; p!=null && seen.Add(p.Id); p=(p as FamilyInstance)?.SuperComponent) ancestors.Add(Id(p.Id));
   var mark = MarkSource.Resolve(bar);
   var parameter = mark.Parameter;
   bool exists = parameter != null;
   string actual = mark.Value;
   var hits = new List<Dictionary<string,object>>(); bool uncertain=false;
   var geometryNotes = new List<string>();
   using(var shape = Geometry.Read(bar)) {
    var box = bar.get_BoundingBox(null);
    if(shape.Failed || shape.Solids.Count==0 || box==null) {
     uncertain=true; geometryNotes.Add("Отсутствует полная твёрдотельная геометрия стержня.");
     geometryNotes.AddRange(shape.Errors);
    } else {
     // Element bounding boxes returned by get_BoundingBox are model-axis aligned.
     using(var outline = new Outline(box.Min,box.Max))
     using(var collector = new FilteredElementCollector(doc)) {
      var candidates = collector.WhereElementIsNotElementType().WherePasses(new ElementMulticategoryFilter(cats))
       .WherePasses(new BoundingBoxIntersectsFilter(outline)).ToElements();
      foreach(var host in candidates) {
       if(host.Id == bar.Id) continue;
       Shape hostShape;
       if(!cache.TryGetValue(host.Id,out hostShape)) cache[host.Id]=hostShape=Geometry.Read(host);
       var hostNotes=new List<string>();
       double volume=0;
       if(hostShape.Failed || hostShape.Solids.Count==0) {
        hostNotes.Add("Геометрия основы недоступна: "+Id(host.Id));
        hostNotes.AddRange(hostShape.Errors.Select(e=>"Основа "+Id(host.Id)+": "+e));
       } else {
        foreach(var solid in hostShape.Solids) {
         int kind; string materialError=null;
         try { kind=Geometry.SolidKind(doc,host,solid); }
         catch(Exception ex) { kind=-1; materialError=ex.Message; }
         if(kind==0) continue;
         try {
          var v=Geometry.Intersection(shape,solid);
          if(v<=Geometry.VolumeTolerance) continue;
          if(kind<0) {
           hostNotes.Add("Неоднозначный материал пересекающегося тела основы "+Id(host.Id));
           if(materialError!=null) hostNotes.Add("Ошибка чтения материала: "+materialError);
          } else volume+=v;
         } catch(Exception ex) {
          hostNotes.Add("Не выполнено пересечение с основой "+Id(host.Id)+": "+ex.GetType().Name+": "+ex.Message);
         }
        }
       }
       if(hostNotes.Count>0) {
        MaterialCodeResult code;
        if(!materialCodes.TryGetValue(host.Id,out code))
         materialCodes[host.Id]=code=MaterialCodeFilter.Read(doc,host);
        // Suppress only this candidate's warnings, not failures of the rebar or other hosts.
        if(code.State==MaterialCodeState.NoMatch) continue;
        uncertain=true;
        geometryNotes.AddRange(hostNotes);
        if(code.State==MaterialCodeState.Concrete)
         geometryNotes.Add("Основа "+Id(host.Id)+": в назначенном материале есть код 20XXXXXXX; предупреждение сохранено.");
        else
         geometryNotes.Add("Основа "+Id(host.Id)+": не удалось достоверно прочитать материалы/коды; предупреждение сохранено. "+String.Join("; ",code.Errors.Distinct()));
       }
       if(volume>Geometry.VolumeTolerance) hits.Add(new Dictionary<string,object>{
        ["host_id"]=Id(host.Id), ["host_name"]=host.Name, ["mark"]=host.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString(),
        ["intersection_volume_ft3"]=volume });
      }
     }
    }
   }
   hits=hits.OrderBy(h=>long.Parse((string)h["host_id"],System.Globalization.CultureInfo.InvariantCulture)).ToList();
   int best=Decision.LargestIntersectionIndex(hits.Select(h=>(double)h["intersection_volume_ft3"]).ToArray());
   var selected=best>=0 && !uncertain ? hits[best] : null;
   var expected=selected==null ? null : selected["mark"] as string;
   var status=Decision.Evaluate(hits.Count,uncertain,exists,actual,expected);
   return new Dictionary<string,object> {
    ["status"]=status, ["element_ids"]=new[]{Id(bar.Id),Id(mark.Element.Id)}.Concat(hits.Select(h=>(string)h["host_id"])).Distinct().ToArray(),
    ["rebar_id"]=Id(bar.Id), ["rebar_name"]=bar.Name, ["parent_id"]=parent==null ? "" : Id(parent.Id),
    ["ancestor_ids"]=ancestors, ["actual_mark"]=actual, ["rebar_mark"]=mark.RebarValue,
    ["mark_source_id"]=Id(mark.Element.Id), ["mark_source_kind"]=mark.Element.Id==bar.Id ? "rebar" : "parent", ["expected_mark"]=expected, ["candidates"]=hits,
    ["selected_host_id"]=selected==null ? null : selected["host_id"],
    ["selected_intersection_volume_ft3"]=selected==null ? null : selected["intersection_volume_ft3"],
    ["host_selection"]="largest_intersection_then_lowest_id",
    ["parameter_read_only"]=parameter?.IsReadOnly, ["geometry_notes"]=geometryNotes,
    ["description"]= status=="mark_mismatch" ? "Марка конструкции не совпадает с маркой IFC-арматуры." :
     status=="match" ? "Марки совпадают." : status=="geometry_uncertain" ? "Сравнение марок не выполнено: "+String.Join("; ",geometryNotes.Distinct()) : "Статус: "+status
   };
  }
 }
}