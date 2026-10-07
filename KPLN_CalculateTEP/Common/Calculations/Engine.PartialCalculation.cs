using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            public static bool MetricNeedsGround(string metric)
            {return metric.StartsWith("Gns") || metric.StartsWith("Np") || OneOf(metric,"GrossAbove","GrossBelow","Storeys","Footprint","ParkingCount");}
            public static bool MetricNeedsZero(string metric)
            {return OneOf(metric,"VolumeAbove","VolumeBelow");}
            private static bool CountMetric(string metric)
            {return OneOf(metric,"ApartmentsCount","ParkingCount","Storeys","Floors");}
            public static bool ParameterAffectsMetric(string key,string metric)
            {
                if(key=="apartment")return metric.StartsWith("Apartments");
                if(key=="parking")return metric=="ParkingCount";
                if(key=="coefficient")return metric=="ApartmentsTotal";
                if(OneOf(key,"function","part"))return !OneOf(metric,"ParkingCount","ApartmentsCount");
                if(OneOf(key,"embedded","standalone"))return OneOf(metric,"NnpEmbedded","NnpSeparate");
                if(OneOf(key,"roof-ratio","service-access","mezzanine-ratio","transition"))return GrossMetric((Indicator)Enum.Parse(typeof(Indicator),metric));
                if(key=="partial-floor")return GrossMetric((Indicator)Enum.Parse(typeof(Indicator),metric))||OneOf(metric,"Storeys","Floors");
                if(key=="$level-top")return MetricNeedsGround(metric);
                if(key.StartsWith("$level-"))return RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),metric))||OneOf(metric,"Storeys","Floors");
                if(OneOf(key,"vertical","height","slope","width","stair-width","floor-offset"))return !CountMetric(metric);
                return true; // An unknown dependency must not silently permit a potentially incomplete total.
            }
            public static bool IssueAffectsMetric(Issue issue,string metric)
            {
                if(!string.IsNullOrEmpty(issue.Metric))return issue.Metric==metric;
                if(issue.Code.StartsWith("AUTO_SHAFT"))return RoomEnvelopeTotal((Indicator)Enum.Parse(typeof(Indicator),metric))||RoomEnvelopePart((Indicator)Enum.Parse(typeof(Indicator),metric))||RoomAreaMetric((Indicator)Enum.Parse(typeof(Indicator),metric));
                if(issue.Code=="PREFLIGHT_PARTIAL")return false;
                if(issue.Code=="ZERO_DATUM")return MetricNeedsZero(metric);
                if(issue.Code=="GROUND_DATUM")return MetricNeedsGround(metric);
                if(issue.Code=="ROOM_AREA_SETTINGS")return RoomAreaMetric((Indicator)Enum.Parse(typeof(Indicator),metric));
                if(issue.Code=="ROOMS_EMPTY")return !metric.StartsWith("Volume")&&!OneOf(metric,"ParkingCount","Footprint","Storeys","Floors");
                if(issue.Code=="APARTMENT_PARAMETER"||issue.Code=="APARTMENT_MULTILEVEL")return metric.StartsWith("Apartments");
                if(issue.InputKind=="room"&&metric=="ParkingCount")return false;
                if(issue.InputKind=="parking")return metric=="ParkingCount";
                if(!string.IsNullOrEmpty(issue.ParameterKey))return ParameterAffectsMetric(issue.ParameterKey,metric);
                if(OneOf(issue.Code,"ROOM_LEVEL","SPATIAL_UNBOUNDED"))return metric!="ParkingCount";
                return true;
            }
            public static bool MetricBlocked(string metric,IEnumerable<Issue> issues)
            {return issues.Any(i=>i.Severity=="Ошибка"&&IssueAffectsMetric(i,metric));}
            // A missing room area taints completeness, but does not prevent processing other rooms/floors.
            // Keep MetricBlocked unchanged for final status, empty results and balance checks.
            public static bool PreflightMetricBlocked(string metric,IEnumerable<Issue> issues)
            {return issues.Any(i=>i.Severity=="Ошибка"&&string.IsNullOrWhiteSpace(i.Element)&&!OneOf(i.Code,"SPATIAL_UNBOUNDED","CATEGORY_FAMILY_EMPTY","SOURCE_UNLOADED")&&IssueAffectsMetric(i,metric));}
            public static string PreflightMessage(IEnumerable<Metric> metrics,IEnumerable<Issue> issues)
            {
                var selected=metrics.Where(m=>m.Enabled).ToList();var errors=issues.ToList();
                int blocked=selected.Count(m=>PreflightMetricBlocked(m.Key,errors));
                int partial=selected.Count(m=>!PreflightMetricBlocked(m.Key,errors)&&MetricBlocked(m.Key,errors));
                if(blocked==0&&partial==0)return "Предварительная проверка завершена. Доступны все выбранные показатели.";
                if(blocked==selected.Count)return "Недостаточно данных для всех выбранных показателей. Причины будут показаны в результате.";
                return "Доступен частичный расчёт: "+(selected.Count-blocked)+" из "+selected.Count+" показателей."+
                    (partial>0?" Для "+partial+" показателей будут рассчитаны доступные данные; пропущенные помещения и этажи указаны в отчёте.":"")+
                    (blocked>0?" Остальные будут отмечены как «Не рассчитано».":"");
            }
            private void ParameterIssue(string code,string message,string key,Record record=null,string source="")
            {
                current.Issues.Add(new Issue{Code=code,Severity="Ошибка",Message=message,ParameterKey=key,InputKind="room",
                    Source=record?.Source.Name??source,Building=record?.Building??"",Element=record==null?"":IDHelper.ElIdValue(record.Element.Id).ToString(),
                    Action="Заполните или исправьте параметр. Независимые показатели рассчитываются автоматически."});
            }
            private void AddSkippedSummary(Metric metric)
            {
                current.Summary.Add(new Summary {Key=metric.Key,Name=metric.Name,Unit=metric.Unit,Method=current.Method,
                    NotCalculated=true,Status="Не рассчитано",Comment=string.Join(" | ",current.Issues.Where(i=>i.Severity=="Ошибка"&&IssueAffectsMetric(i,metric.Key)).Select(i=>i.Code+": "+i.Message).Distinct().Take(3))});
            }
        }
    }
}
