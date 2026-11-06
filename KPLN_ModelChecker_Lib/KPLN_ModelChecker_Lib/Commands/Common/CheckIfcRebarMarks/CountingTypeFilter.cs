using System;
using System.Globalization;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckIfcRebarMarks {
 internal static class CountingTypeFilter {
  internal static bool Exclude(Element rebar,out string warning) {
   warning=null;
   try {
    var own=Read(rebar,out warning);
    if(own.HasValue || warning!=null) return own.GetValueOrDefault();
    var typeId=rebar.GetTypeId();
    if(typeId==null || typeId==ElementId.InvalidElementId) return false;
    var type=rebar.Document.GetElement(typeId);
    if(type==null) { warning="Тип стержня не найден; стержень оставлен в проверке."; return false; }
    return Read(type,out warning).GetValueOrDefault();
   } catch(Exception ex) {
    warning="Не удалось прочитать Орг.ИзделиеТипПодсчета; стержень оставлен в проверке: "+ex.Message;
    return false;
   }
  }
  // null means absent/unset, so the caller may look on the type.
  // false with a warning means unreadable/ambiguous, not a reason to exclude.
  private static bool? Read(Element element,out string warning) {
   warning=null;
   var parameters=element.GetParameters("Орг.ИзделиеТипПодсчета");
   if(parameters.Count==0) return null;
   if(parameters.Count!=1) { warning="Неоднозначный Орг.ИзделиеТипПодсчета; стержень оставлен в проверке."; return false; }
   var p=parameters[0];
   if(!p.HasValue) return null;
   if(p.StorageType==StorageType.Integer) return p.AsInteger()!=1;
   if(p.StorageType==StorageType.Double) {
    var value=p.AsDouble();
    if(double.IsNaN(value)||double.IsInfinity(value)) { warning="Некорректное значение типа подсчёта; стержень оставлен в проверке."; return false; }
    return value!=1.0;
   }
   if(p.StorageType==StorageType.String) {
    var text=p.AsString();
    if(String.IsNullOrWhiteSpace(text)) return null;
    double value;
    if(double.TryParse(text.Trim(),NumberStyles.Float,CultureInfo.InvariantCulture,out value) ||
       double.TryParse(text.Trim(),NumberStyles.Float,CultureInfo.CurrentCulture,out value))
     return value!=1.0;
    return true;
   }
   warning="Неподдерживаемый тип Орг.ИзделиеТипПодсчета; стержень оставлен в проверке.";
   return false;
  }
 }
}
