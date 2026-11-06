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
    public sealed class CheckRebarConcreteClass : AbstrCheck
    {
        public CheckRebarConcreteClass() : base()
        {
            PluginName = "КР: Проверка класса бетона арматуры";
            ESEntity = new ExtensibleStorageEntity(PluginName, "KPLN_CheckRebarConcreteClass",
                new Guid("aa6b1143-1839-4211-9376-bb78bed8bbb2"), new Guid("971826f7-7013-4057-8ac2-7c80b4394ebf"));
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
            int closed = doc.IsWorkshared ? new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).Count(w => !w.IsOpen) : 0;
            var bars = new List<Element>();
            using (var hosts = new Hosts(doc))
            {
                bool fast = Rules.FastMode(hosts.Classes, hosts.InventoryWarnings.Count == 0, closed);
                _checkerEntitiesCollHeap.AddRange(hosts.InventoryWarnings);
                foreach (var bar in elemColl)
                {
                    string flagError;
                    if (!Parameters.RunningMeters(bar, out flagError))
                    {
                        if (flagError != null)
                            _checkerEntitiesCollHeap.Add(new CheckerEntity(bar, "Не проверено",
                                "Не прочитан Рзм.ПогМетрыВкл. " + flagError, string.Empty).Set_Status(ErrorStatus.Warning));
                        continue;
                    }
                    string countingError;
                    if (CountingTypeFilter.Exclude(bar, out countingError)) continue;
                    if (countingError != null)
                        _checkerEntitiesCollHeap.Add(new CheckerEntity(bar, "Ограничение проверки",
                            "Не удалось однозначно прочитать Орг.ИзделиеТипПодсчета. " + countingError, string.Empty).Set_Status(ErrorStatus.Warning));
                    bars.Add(bar);
                }
                foreach (var bar in bars)
                {
                    try
                    {
                        var entity = Inspect(bar, hosts, fast);
                        if (entity != null) _checkerEntitiesCollHeap.Add(entity);
                    }
                    catch (Exception ex)
                    {
                        _checkerEntitiesCollHeap.Add(new CheckerEntity(bar, "Не проверено",
                            "Ошибка обработки элемента: " + ex.GetType().Name + ": " + ex.Message, string.Empty).Set_Status(ErrorStatus.Warning));
                    }
                }
            }
            if (bars.Count == 0 && _checkerEntitiesCollHeap.Count == 0)
            {
                _checkerEntitiesCollHeap.Add(new CheckerEntity("Отсутствуют элементы для проверки",
                    "Нет подходящей погонной арматуры после отбора и исключений. Проверяется текущая модель без связей.", string.Empty));
                return CheckResultStatus.NoItemsToCheck;
            }
            return CheckResultStatus.Succeeded;
        }

  private CheckerEntity Inspect(Element bar,Hosts hosts,bool fast) {
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
   if (status == "match") return null;
   bool error = new[]{"class_mismatch","rebar_class_empty","rebar_class_invalid"}.Contains(status);
   var info = new StringBuilder();
   info.AppendLine("Фактический класс: «" + actual.Text + "»");
   info.AppendLine("Ожидаемый класс: " + host.Class);
   info.AppendLine("ID выбранной основы: " + Id(host.Id));
   var entity = new CheckerEntity(bar, error ? "Ошибка класса бетона" : "Требует разбора",
    messages.ContainsKey(status) ? messages[status] : status, info.ToString());
   return error ? entity.Set_CanApprovedAndESData(ESEntity) : entity.Set_Status(ErrorStatus.Warning);
  }
    }
}
