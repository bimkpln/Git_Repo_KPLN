using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CalculateTEP.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_CalculateTEP.Common.Geometry.TepClipper.Clipper;
using CP = KPLN_CalculateTEP.Common.Geometry.TepClipper.IntPoint;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;
using PT = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyType;
using PF = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyFillType;

using TepClipper = KPLN_CalculateTEP.Common.Geometry.TepClipper;
using KPLN_CalculateTEP.Common.Methodologies;
namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private sealed class Timing {internal long Ticks,Calls;}
            private sealed class SlowOperation
            {internal string Stage,Metric,Source,Building,Element;internal double Seconds;}
            private T Measure<T>(string stage,Record record,string metric,Func<T> action)
            {
                long start=System.Diagnostics.Stopwatch.GetTimestamp();
                try{return action();}
                finally
                {
                    long ticks=System.Diagnostics.Stopwatch.GetTimestamp()-start;Timing timing;
                    if(!timings.TryGetValue(stage,out timing))timings[stage]=timing=new Timing();
                    timing.Ticks+=ticks;timing.Calls++;
                    double seconds=(double)ticks/System.Diagnostics.Stopwatch.Frequency;
                    if(seconds>=.5)
                    {
                        slowOperations.Add(new SlowOperation{Stage=stage,Metric=metric,Source=record?.Source.Name,Building=record?.Building,
                            Element=record==null?null:IDHelper.ElIdValue(record.Element.Id).ToString(),Seconds=seconds});
                        if(slowOperations.Count>20)slowOperations.Remove(slowOperations.OrderBy(x=>x.Seconds).First());
                    }
                }
            }
            private void WriteTimings()
            {
                foreach(var pair in timings.OrderByDescending(p=>p.Value.Ticks))
                    current.Issue("PERF_STAGE","Информация",pair.Key+": "+((double)pair.Value.Ticks/System.Diagnostics.Stopwatch.Frequency).ToString("0.###")+" с; операций: "+pair.Value.Calls,
                        action:"Времена этапов включают вложенные операции и не суммируются. Используйте для сравнения повторных запусков на одной модели.");
                current.Issue("PERF_CACHE","Информация","Повторно использовано наборов объёмных фрагментов: "+volumeCacheHits+"; локальных объёмов помещений / пространств: "+spatialCacheHits,
                    action:"Кэш действует только внутри одного запуска; набор фрагментов повторно используется при совпадении входных объектов и их расчётных назначений.");
                current.Issue("PERF_BOOLEAN","Информация","Резервная обработка объёмов: без объёмного пересечения - "+booleanTouchSkips+"; разбиением на связные части - "+booleanSplitRecoveries+"; через пересечение A ∩ B - "+booleanIntersectionRecoveries+"; в локальных координатах - "+booleanNormalizedRecoveries,
                    action:"Резервные операции запускаются после отказа прямого вычитания. Необработанные фрагменты сохраняют ошибку VOLUME_BOOLEAN.");
                current.Issue("PERF_PLANAR","Информация","Вычитаний объёмов по высотным участкам без Boolean Revit: "+planarVolumeCuts);
                foreach(var slow in slowOperations.OrderByDescending(x=>x.Seconds))
                    current.Issue("PERF_SLOW","Информация",slow.Stage+": "+slow.Seconds.ToString("0.###")+" с",slow.Metric??"",slow.Source??"",slow.Building??"",slow.Element??"",
                        "Одна из 20 наиболее длительных операций от 0,5 с. Перейдите к объекту для проверки сложности геометрии.");
            }
            private void Notice(string code,string severity,string message,Record r=null,string metric="",string action="Уточните классификацию или исходную геометрию.")
            {
                string key=code+"|"+metric+"|"+(r==null?"":r.Key)+"|"+message;if(!notices.Add(key))return;
                if(r?.Source.Mode=="reference"&&severity=="Ошибка")severity="Предупреждение";
                current.Issue(code,severity,message,metric,r?.Source.Name??"",r?.Building??"",r==null?"":IDHelper.ElIdValue(r.Element.Id).ToString(),action);
            }
            private void RefreshStatuses()
            {
                foreach(var s in current.Summary)
                {
                    var issues=current.Issues.Where(x=>IssueAffectsMetric(x,s.Key)&&!OneOf(x.Code,"SAVE_FAILED","SCHEDULE_FAILED","GRAPHICS_FAILED","GRAPHIC_ELEMENT","GRAPHICS_NO_GEOMETRY","CLEANUP_DEPENDENCY","REVIT_TRANSACTION")).ToList();
                    if(s.NotCalculated)s.Status="Не рассчитано";
                    else if(issues.Any(x=>x.Severity=="Ошибка"))s.Status="Неполный результат";
                    else if(s.Status!="Нет данных"&&issues.Any(x=>x.Severity=="Предупреждение"))s.Status="Требует проверки";
                    s.Comment=s.Key=="Storeys"||s.Key=="Floors"?"Итог проекта - максимум по корпусам и секциям. ":"";
                    s.Comment+=string.Join(" | ",issues.Where(x=>x.Severity!="Информация").Select(x=>x.Code+": "+x.Message).Distinct().Take(3));
                    s.Value=Math.Abs(s.Value)<1e-10?0:s.Value;
                }
            }
            private void CheckBalances()
            {
                var balances=new[]{new[]{"Gns","GnsResidential","GnsNonresidential"},new[]{"GnsResidential","GnsLivingPart","GnsNonlivingPart"},new[]{"Gross","GrossAbove","GrossBelow"},new[]{"Volume","VolumeAbove","VolumeBelow"},new[]{"Np","NpResidential","NpNonresidential"}};
                foreach(var keys in balances)
                {
                    var all=keys.Select(k=>current.Summary.FirstOrDefault(s=>s.Key==k)).ToList();if(all.Any(x=>x==null||x.NotCalculated)||keys.Any(k=>MetricBlocked(k,current.Issues)))continue;
                    if(Math.Abs(all[0].Value-all[1].Value-all[2].Value)>.01)
                        foreach(var key in keys)current.Issue("BALANCE","Ошибка","Не выполнено равенство: "+keys[0]+" = "+keys[1]+" + "+keys[2]+". Проверьте правила отдельных показателей и классификацию.",key);
                }
                var full=current.Summary.FirstOrDefault(s=>s.Key=="ApartmentsTotal");var heated=current.Summary.FirstOrDefault(s=>s.Key=="ApartmentsHeated");
                if(full!=null&&heated!=null&&!full.NotCalculated&&!heated.NotCalculated&&!MetricBlocked(full.Key,current.Issues)&&!MetricBlocked(heated.Key,current.Issues)&&full.Value+.01<heated.Value)current.Issue("APARTMENT_BALANCE","Ошибка","Площадь с летними помещениями меньше площади без них.","ApartmentsTotal");
            }
        }
    }
}
