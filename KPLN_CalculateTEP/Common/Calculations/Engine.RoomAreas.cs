using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private static bool RoomAreaMetric(Indicator metric)
            {return OneOf(metric.ToString(),"ApartmentsTotal","ApartmentsHeated","PublicRooms","PublicCalculated");}
            public static bool IsRoomAreaMaskRole(string role)
            {return OneOf(role,"multilight","stair-gap","opening","shaft","engineering-shaft","stove","decoration");}
            private bool RoomAreaMaskRequired(Record r,Indicator metric)
            {
                if(r.Override==false)return true;
                return IsRoomAreaMaskRole(r.Role);
            }
            private void CalculateRoomAreas(List<Record> records,Indicator metric)
            {
                Func<Record,string> floorKey=r=>r.Building+"|"+Math.Round(r.Z,6).ToString("R",System.Globalization.CultureInfo.InvariantCulture);
                var rooms=records.Where(r=>IsAreaInput(r.Element)&&r.Level!=null&&r.Level.Include&&r.Source.Mode!="reference").ToList();
                var included=new List<Record>();var invalid=new HashSet<string>();
                foreach(var r in rooms)
                {
                    try{if(Eligible(r,metric))included.Add(r);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){invalid.Add(floorKey(r));Notice("ROOM_AREA_INPUT","Ошибка",ex.Message,r,metric.ToString());}
                }
                if(included.Count==0)return;
                var targetFloors=new HashSet<string>(included.Select(floorKey));
                var masks=new Dictionary<string,PlanarRegion>();var previous=new Dictionary<string,PlanarRegion>();
                foreach(var r in records.Where(r=>r.Level!=null&&r.Level.Include&&r.Source.Mode!="reference"&&targetFloors.Contains(floorKey(r))))
                {
                    try
                    {
                        bool room=IsAreaInput(r.Element);
                        if(!room&&!IsRoomAreaMaskRole(r.Role))continue;
                        bool eligible=Eligible(r,metric);
                        if(room&&(eligible||!RoomAreaMaskRequired(r,metric)))continue;
                        // Keep separately modelled shafts/openings and the same vertical-space rules.
                        var region=room?EligibleRoomRegion(r,metric):ReviewRegion("room-mask/"+metric+"/"+r.Key,r,"Исключение - "+RoleLabel(r.Role),r.Z,()=>RegionOfPlan(Plan(r,metric)));
                        if(region.IsEmpty)continue;
                        if(!room&&eligible&&!VerticalExclusionArea(r,metric,records,region.Area))continue;
                        string key=floorKey(r);PlanarRegion existing;
                        masks[key]=RegionUnion(masks.TryGetValue(key,out existing)?existing:PlanarRegion.Empty,region);
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){invalid.Add(floorKey(r));Notice("ROOM_AREA_MASK","Ошибка",ex.Message,r,metric.ToString());}
                }
                var reported=new HashSet<string>();
                foreach(var r in included)
                {
                    string key=floorKey(r);
                    if(invalid.Contains(key))
                    {
                        if(reported.Add(key))Notice("ROOM_AREA_FLOOR_SKIPPED","Ошибка","Этаж «"+r.Level.Name+"» пропущен: состав исключений не подтверждён. Остальные этажи рассчитываются.",r,metric.ToString());
                        continue;
                    }
                    Progress("2D-площадь помещения: "+MetricLabels[metric.ToString()]+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    try
                    {
                        if(IsClassifiedFamily(r.Element)&&UsesFamilyParameterArea(r.Element))
                        {
                            double area=FamilyParameterArea(r);
                            if(area<=1e-9){Notice("ZERO_AREA","Предупреждение","Нулевая площадь экземпляра: объект пропущен.",r,metric.ToString());continue;}
                            if(masks.ContainsKey(key)||IsRoomAreaMaskRole(r.Role)||OneOf(r.Role,"under-stair","mezzanine")||NeedsCeilingCheck(r,metric)&&HasSlopedCeiling(r))
                                throw new InvalidOperationException("Для геометрических исключений и высотных проверок недостаточно готовой площади. Выберите геометрию экземпляра или привяжите расчётную область.");
                            current.Details.Add(Row(r,metric,area*.09290304,Factor(r,metric),false,"Параметр площади экземпляра: "+FamilyAreaParameterName(r.Element)));
                            Notice("FAMILY_AREA_PARAMETER","Предупреждение","Площадь экземпляра взята из параметра. Пространственное пересечение с помещениями не проверено; исключите дублирование в классификации.",r,metric.ToString());
                            continue;
                        }
                        var region=EligibleRoomRegion(r,metric);if(region.IsEmpty){Notice("ZERO_AREA","Предупреждение","Пустая расчётная область: объект пропущен.",r,metric.ToString());continue;}
                        bool excluded=VerticalExclusionArea(r,metric,records,region.Area);
                        if(!excluded)
                        {
                            PlanarRegion mask,used;
                            if(masks.TryGetValue(key,out mask))region=RegionDifference(region,mask);
                            if(region.IsEmpty)continue;
                            if(previous.TryGetValue(key,out used)&&RegionIntersection(used,region).Area*.09290304>.005)
                                throw new InvalidOperationException("Помещение пересекается с ранее учтённым более чем на 0,005 м². Дубликат не включён в сумму; результат неполный.");
                            previous[key]=RegionUnion(used??PlanarRegion.Empty,region);
                        }
                        var row=Row(r,metric,region.Area*.09290304,Factor(r,metric),excluded,
                            (editedObjectOverlays.TryGetValue(ObjectOverlayKey(r),out var edited)&&edited!=null?"Исправленная цветовая область объекта":"Граница расчётной 2D-области нормативного обмера")+"; 2D-исключения; коэффициент применяется один раз");
                        row.MeasurementElevation=edited!=null?r.Z:RoomFloorElevation(r)+RoomMeasurementHeight("net",metric.ToString(),r.Role)/.3048;
                        current.Details.Add(row);
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){Notice("ROOM_AREA","Ошибка",ex.Message,r,metric.ToString());}
                }
                Notice("ROOM_PLANAR_AREAS","Информация","Площади помещений получены независимо от наружной оболочки здания и наличия отдельного перекрытия. Использованы нормативные сечения Room, плоские маски и проверки пересечений.",included[0],metric.ToString());
            }
        }
    }
}
