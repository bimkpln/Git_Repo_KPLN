using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KPLN_ModelChecker_Lib.Commands.Common.CheckIfcRebarMarks;
using KPLN_ModelChecker_Lib.Commands.Common.CheckRebarConcreteClass;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckIfcRebarMarks : AbstrCheck
    {
        public CheckIfcRebarMarks() : base()
        {
            PluginName = "КР: Проверка марок IFC-арматуры";
            ESEntity = new ExtensibleStorageEntity(PluginName, "KPLN_CheckIfcRebarMarks",
                new Guid("15710cbc-a9bb-4bc9-bb19-ffc855614561"), new Guid("3d54de59-b5e8-45ff-9b9a-67b620cb00cc"));
        }

        public override Element[] GetElemsToCheck()
        {
            return new FilteredElementCollector(CheckDocument)
                .OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToArray();
        }

        internal static string Id(ElementId id)
        {
            if (id == null) return "";
#if Debug2024 || Revit2024
            return id.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
#else
            return id.IntegerValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
#endif
        }

        private protected override CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl)
        {
            _checkerEntitiesCollHeap.Clear();
            var doc = CheckDocument;
            var bars = new List<Element>();
            int unidentified = 0;
            foreach (var e in elemColl)
            {
                var flag = UniqueParameter(doc.GetElement(e.GetTypeId()), "Арм.ВыполненаСемейством");
                if (flag == null || flag.StorageType != StorageType.Integer) { unidentified++; continue; }
                if (flag.AsInteger() != 1) continue;
                string warning;
                if (CountingTypeFilter.Exclude(e, out warning)) continue;
                if (warning != null)
                    _checkerEntitiesCollHeap.Add(new CheckerEntity(e, "Ограничение проверки",
                        "Не удалось однозначно прочитать Орг.ИзделиеТипПодсчета. " + warning, string.Empty).Set_Status(ErrorStatus.Warning));
                bars.Add(e);
            }
            var categories = new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_StairsRuns, BuiltInCategory.OST_StairsLandings };
            var cache = new Dictionary<ElementId, Shape>();
            var materialCodes = new Dictionary<ElementId, MaterialCodeResult>();
            try
            {
                foreach (var bar in bars)
                {
                    try
                    {
                        var entity = Inspect(doc, bar, categories, cache, materialCodes);
                        if (entity != null) _checkerEntitiesCollHeap.Add(entity);
                    }
                    catch (Exception ex)
                    {
                        _checkerEntitiesCollHeap.Add(new CheckerEntity(bar, "Не проверено",
                            "Ошибка обработки элемента: " + ex.GetType().Name + ": " + ex.Message, string.Empty).Set_Status(ErrorStatus.Warning));
                    }
                }
            }
            finally { foreach (var shape in cache.Values) shape.Dispose(); }
            if (unidentified > 0)
                _checkerEntitiesCollHeap.Add(new CheckerEntity("Ограничение проверки", "Не определён признак IFC у " + unidentified + " элементов арматуры. Нет однозначного параметра типа Арм.ВыполненаСемейством.",
                    string.Empty).Set_Status(ErrorStatus.Warning));
            if (bars.Count == 0 && _checkerEntitiesCollHeap.Count == 0)
            {
                _checkerEntitiesCollHeap.Add(new CheckerEntity("Отсутствуют элементы для проверки",
                    "Нет подходящей IFC-арматуры после отбора и исключений. Проверяется текущая модель без связей.", string.Empty));
                return CheckResultStatus.NoItemsToCheck;
            }
            return CheckResultStatus.Succeeded;
        }

        internal static Parameter UniqueParameter(Element element, string name)
        {
            if (element == null) return null;
            var parameters = element.GetParameters(name);
            return parameters.Count == 1 ? parameters[0] : null;
        }

  private CheckerEntity Inspect(Document doc, Element bar, List<BuiltInCategory> cats, Dictionary<ElementId,Shape> cache, Dictionary<ElementId,MaterialCodeResult> materialCodes) {
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
   if (status == "match") return null;
   bool error = new[]{"mark_mismatch","rebar_mark_empty","host_mark_empty","parameter_missing"}.Contains(status);
   var messages = new Dictionary<string,string> {
    {"mark_mismatch", "Марки конструкции и IFC-арматуры не совпадают."},
    {"rebar_mark_empty", "Марка арматуры не заполнена."},
    {"host_mark_empty", "Марка бетонной основы не заполнена."},
    {"parameter_missing", "Параметр марки недоступен или неоднозначен."},
    {"geometry_uncertain", "Не удалось достоверно определить основу по геометрии/материалам."},
    {"no_host", "Бетонная основа не найдена."},
    {"host_not_found", "Бетонная основа не найдена."}
   };
   var info = new StringBuilder();
   info.AppendLine("Марка основы арматуры: «" + actual + "»");
   info.AppendLine("Марка ожидаемой основы: «" + expected + "»");
   info.AppendLine("ID ожидаемой основы: " + (selected == null ? "не определена" : selected["host_id"]));
   var entity = new CheckerEntity(bar, error ? "Ошибка марки IFC-арматуры" : "Требует разбора",
    messages.ContainsKey(status) ? messages[status] : status, info.ToString());
   return error ? entity.Set_CanApprovedAndESData(ESEntity) : entity.Set_Status(ErrorStatus.Warning);
  }
    }
}
