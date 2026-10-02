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
            public Run Calculate(Action<string> progress=null,bool? createViews=null)
            {
                reportProgress=progress;var previous=Last;
                using(var group=new TransactionGroup(doc,"ТЭП: расчёт и проверочные виды"))
                {
                    group.Start();
                    try {var run=CalculateCore(progress,createViews);Progress("Завершение расчёта...");if(group.Assimilate()!=TransactionStatus.Committed)throw new InvalidOperationException("Revit отменил группу транзакций расчёта.");Last=run;return run;}
                    catch {if(group.GetStatus()==TransactionStatus.Started)group.RollBack();Last=previous;throw;}
                    finally {createViewsForRun=null;planarBodies=null;planarCheckpoint=null;reportProgress=null;parameterCache.Clear();phaseCache.Clear();ClearVolumeCaches();floorSurfaces.Clear();floorWallFaces.Clear();floorFaceHeights.Clear();nativeWallLayers.Clear();roomHeightSections.Clear();}
                }
            }
            private Run CalculateCore(Action<string> progress,bool? createViews)
            {
                planarBodies=new Dictionary<Solid,Tuple<LayeredBody,string>>();progressMessage="Подготовка расчёта контуров...";planarCheckpoint=()=>reportProgress?.Invoke(progressMessage);planarVolumeCuts=0;
                parameterCache.Clear();phaseCache.Clear();ClearVolumeCaches();timings.Clear();slowOperations.Clear();volumeCacheHits=0;spatialCacheHits=0;booleanTouchSkips=0;booleanSplitRecoveries=0;booleanIntersectionRecoveries=0;booleanNormalizedRecoveries=0;
                PrepareRoomWorkflow();SnapshotSources();Normalize(Config);notices.Clear();shapes.Clear();appliedCorrections.Clear();roomFloorContours.Clear();roomFloorFailures.Clear();invalidRoomFloors.Clear();floorSurfaces.Clear();floorWallFaces.Clear();floorFaceHeights.Clear();nativeWallLayers.Clear();roomHeightSections.Clear();
                if(createViews.HasValue)Config.CreateViews=createViews.Value;
                createViewsForRun=Config.CreateViews;
                current=new Run{Author=app.Application.Username,Method=Choices("method").First(x=>x.Key==Config.Method).Label};
                current.CreateViews=ViewsEnabled;
                current.Issues.AddRange(startupIssues.Where(i=>string.IsNullOrEmpty(i.Source)||Sources.Any(s=>s.Name==i.Source&&s.Mode!="exclude")));
                foreach(var c in Config.Corrections)
                {if(string.IsNullOrWhiteSpace(c.Author))c.Author=current.Author;if(string.IsNullOrWhiteSpace(c.Date))c.Date=current.Date;}
                current.Configuration=Serialize(Config);
                ValidateDatums();
                var records=PreflightRecords();
                var preflightIssues=current.Issues.ToList();
                var blockedMetrics=new HashSet<string>(Config.Metrics.Where(m=>m.Enabled&&MetricBlocked(m.Key,preflightIssues)).Select(m=>m.Key));
                if(blockedMetrics.Count>0)
                {
                    string message=PreflightMessage(Config.Metrics,preflightIssues);
                    current.Issue("PREFLIGHT_PARTIAL","Предупреждение",message);
                    progress?.Invoke(message+" Выполняю доступные показатели.");
                }
                current.Issue("ROOM_WORKFLOW","Информация","ГНС: внешнее временное помещение с резервным сечением стен; площади квартир и помещений: Room.Area с коэффициентами, особыми нормативными проверками и контролем пересечений в 2D. Объёмы сохраняют геометрический расчёт.");
                var invalidInputs=new HashSet<double>(invalidRoomFloors);
                foreach(var metric in Config.Metrics.Where(x=>x.Enabled))
                {
                    if(blockedMetrics.Contains(metric.Key)){AddSkippedSummary(metric);continue;}
                    invalidRoomFloors.Clear();invalidRoomFloors.UnionWith(invalidInputs);
                    progress?.Invoke("Расчёт: "+metric.Name);var indicator=(Indicator)Enum.Parse(typeof(Indicator),metric.Key);
                    try
                    {
                        var adjusted=new List<Record>();
                        foreach(var original in records)
                        {
                            Progress("Правила: "+metric.Name+" - "+(adjusted.Count+1)+" / "+records.Count);
                            var r=original.Copy();try{ApplyRules(r,metric.Key);adjusted.Add(r);}
                            catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){MarkInvalidRoomFloor(r.Source,r.Element);Notice("RULE_CONFLICT","Ошибка",ex.Message,r,metric.Key);}
                        }
                        adjusted=Representations(adjusted,metric);
                        foreach(var group in adjusted.GroupBy(r=>r.Building))
                        {bool single=group.Where(r=>r.Level!=null&&r.Level.Include&&r.Role!="mezzanine").Select(r=>Math.Round(r.Z,6)).Distinct().Count()==1;foreach(var r in group)r.SingleStorey=single;}
                        if(RoomEnvelopeTotal(indicator))CalculateRoomEnvelopeAreas(adjusted,indicator);
                        else if(metric.ContourMode!="current")CalculateContours(adjusted,indicator,metric);
                        else if(indicator==Indicator.Storeys||indicator==Indicator.Floors)CalculateFloors(adjusted,indicator);
                        else if(indicator==Indicator.ApartmentsCount||indicator==Indicator.ParkingCount)CalculateCounts(adjusted,indicator);
                        else if(indicator==Indicator.Volume||indicator==Indicator.VolumeAbove||indicator==Indicator.VolumeBelow)CalculateVolumes(adjusted,indicator);
                        else if(indicator==Indicator.Footprint)CalculateFootprints(adjusted);
                        else if(RoomAreaMetric(indicator))CalculateRoomAreas(adjusted,indicator);
                        else CalculateAreas(adjusted,indicator);
                        ApplyDeltas(indicator);
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){current.Issue("METRIC_FAILED","Ошибка",ex.Message,metric.Key);}
                    current.Issue("CALCULATION_PATH","Информация","Геометрия "+GeometryVersion+"; источник контура: "+Choices("contour").First(c=>c.Key==metric.ContourMode).Label+
                        ". Плоские операции: сетка 0,001 мм; отклонение аппроксимации дуг, эллипсов и сплайнов не более 0,1 мм. Непризматические тела сохраняют обработку Revit.",metric.Key);
                    Summarize(metric);
                }
                foreach(var c in Config.Corrections.Where(c=>!appliedCorrections.Contains(c)))
                {
                    if(c.Metric!="all"&&!Config.Metrics.Any(m=>m.Key==c.Metric&&m.Enabled))continue;
                    current.Issue("CORRECTION_NOT_APPLIED","Ошибка","Корректировка не применена: объект / источник не найден либо действие неприменимо. "+c.Reason,c.Metric=="all"?"":c.Metric,c.Source,element:c.Element,
                        action:"Проверьте источник, ID, выбранный показатель и действие. Для числовой дельты выберите конкретный показатель.");
                }
                CheckBalances();AddErrorContours();current.Corrections=Deserialize<List<Correction>>(Serialize(Config.Corrections.ToList()));
                current.Configuration=Serialize(Config);Last=current;
                DisposeSpatialCalculators();
                if(ViewsEnabled) {progress?.Invoke("Построение проверочных видов...");Measure("Создание проверочной графики",null,"",()=>{CreateGraphics(current);return true;});}
                RefreshStatuses();
                if(ViewsEnabled&&Config.CreateSchedule){progress?.Invoke("Создание сводной спецификации...");CreateSchedule(current);}
                RefreshStatuses();
                WriteTimings();
                try{if(!settingsLoadFailed)SaveSettings();Write("run",current);}catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){current.Issue("SAVE_FAILED","Ошибка","Расчёт выполнен, но запись в RVT не удалась: "+ex.Message);RefreshStatuses();}
                return current;
            }
            private Detail Row(Record r,Indicator metric,double raw,double factor=1,bool excluded=false,string reason="",Solid shape=null,bool volume=false)
            {
                double? measurement=null;
                if(!volume&&shape!=null&&r.Element is Room&&metric!=Indicator.Footprint&&metric!=Indicator.ApartmentsCount&&metric!=Indicator.ParkingCount)
                {
                    string boundary=(int)metric<=4?"outer":GrossMetric(metric)?"gross":"net";
                    measurement=RoomFloorElevation(r)+RoomMeasurementHeight(boundary,metric.ToString(),r.Role)/.3048;
                    reason+="; отметка обмера "+(measurement.Value*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м";
                }
                return new Detail{Metric=metric.ToString(),SourceKey=r.Source.Key,Source=r.Source.Name,Building=r.Building,Section=r.Section,
                    Level=r.Level?.Name??"Без уровня",Elevation=r.Z,MeasurementElevation=measurement,Element=IDHelper.ElIdValue(r.Element.Id).ToString(),UniqueId=r.Element.UniqueId,
                    Purpose=r.Role,Apartment=r.Apartment,Profile=r.Profile,Method=current.Method,Unit=Config.Metrics.First(m=>m.Key==metric.ToString()).Unit,
                    Raw=raw,Factor=factor,Value=excluded?0:raw*factor,Excluded=excluded,Manual=r.Manual,Reason=reason,Shape=shape,VolumeShape=volume};
            }
            private void CalculateAreas(List<Record> records,Indicator metric)
            {
                if((metric==Indicator.ApartmentsTotal||metric==Indicator.ApartmentsHeated)&&string.IsNullOrWhiteSpace(Config.Parameter("apartment")))
                    throw new InvalidOperationException("Для площади квартир задайте параметр ID / номера квартиры.");
                var masks=new Dictionary<string,PlanIndex>();var used=new Dictionary<string,PlanIndex>();
                int processed=0;
                var pending=new List<Tuple<Record,Solid,double,bool,string>>();
                foreach(var r in records)
                {
                    Progress("Контуры: "+MetricLabels[metric.ToString()]+" - "+(++processed)+" / "+records.Count+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    // Construction solids are used by envelope/volume routines, never as room floor areas.
                    if(OneOf(r.Role,"structure","envelope","footprint","underground-footprint")&&!r.Override.HasValue)continue;
                    if(!(r.Element is Room)&&!OneOf(r.Role,"multilight","stair-gap","opening","shaft","engineering-shaft","stove","decoration"))continue;
                    if(r.Level==null){Notice("LEVEL_MISSING","Ошибка","Не определён уровень расчётного объекта.",r,metric.ToString());continue;}
                    if(!r.Level.Include)continue;
                    try
                    {
                        bool include=Eligible(r,metric);
                        var shape=Plan(r,metric);if(shape==null||shape.Volume<1e-9)continue;
                        if(RoomEnvelopePart(metric)&&r.Element is Room)shape=ClipRoomPart(shape,r,records,metric);
                        if(include&&VerticalExclusion(r,metric,records,shape))include=false;
                        string reason=include?"Включено по профилю и классификации":"Исключено по профилю, классификации или этажу";
                        if(r.Override.HasValue)reason="Ручное правило: "+(include?"включить":"исключить");
                        bool reference=r.Source.Mode=="reference";
                        string key=r.Building+"|"+r.Section+"|"+Math.Round(r.Z,6).ToString("R",CultureInfo.InvariantCulture);
                        if(!include&&!reference&&(r.Element is SpatialElement||r.Override==false||OneOf(r.Role,"multilight","stair-gap","opening","shaft","stove","decoration")))
                        {PlanIndex index;if(!masks.TryGetValue(key,out index))masks[key]=index=new PlanIndex();index.Add(shape,r);}
                        pending.Add(Tuple.Create(r,shape,Factor(r,metric),include&&!reference,reference?"Источник только для графической сверки":reason));
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("AREA_INPUT","Ошибка",ex.Message,r,metric.ToString());}
                }
                processed=0;
                foreach(var item in pending.OrderBy(x=>x.Item1.Key,StringComparer.Ordinal))
                {
                    Progress("Пересечения: "+MetricLabels[metric.ToString()]+" - "+(++processed)+" / "+pending.Count+"; ID "+IDHelper.ElIdValue(item.Item1.Element.Id));
                    var r=item.Item1;var shape=item.Item2;string key=r.Building+"|"+r.Section+"|"+Math.Round(r.Z,6).ToString("R",CultureInfo.InvariantCulture);
                    if(!item.Item4){current.Details.Add(Row(r,metric,shape.Volume*.09290304,item.Item3,true,item.Item5,shape));continue;}
                    try
                    {
                        PlanIndex index;
                        if(masks.TryGetValue(key,out index))foreach(var mask in index.Query(shape))
                        {Progress("Вычитание исключений: "+metric+"; ID "+IDHelper.ElIdValue(r.Element.Id));shape=Subtract(shape,mask);if(shape==null||shape.Volume<1e-9)break;}
                        var unique=shape;
                        if(used.TryGetValue(key,out index))foreach(var previous in index.Query(shape))
                        {Progress("Проверка пересечений: "+metric+"; ID "+IDHelper.ElIdValue(r.Element.Id));unique=Subtract(unique,previous);if(unique==null||unique.Volume<1e-9)break;}
                        double duplicate=Math.Max(0,((shape?.Volume??0)-(unique?.Volume??0))*.09290304);
                        if(duplicate>.005)Notice("AREA_OVERLAP","Предупреждение","Пересечение расчётных контуров "+duplicate.ToString("0.###")+" м² учтено один раз. Проверьте дубликаты и принадлежность квартиры / функции.",r,metric.ToString());
                        if(!used.TryGetValue(key,out index))used[key]=index=new PlanIndex();
                        index.Add(unique,r);
                        double raw=unique==null?0:unique.Volume*.09290304;
                        if(!IsPublic(r)&&r.Role=="transition"&&GrossMetric(metric))
                        {
                            var buildings=(Mapped(r.Element,"transition")??"").Split(';').Select(x=>x.Trim()).Where(x=>x.Length>0).Distinct().ToList();
                            if(buildings.Count<2)throw new InvalidOperationException("Для перехода укажите минимум два соединяемых корпуса через ;.");
                            foreach(var building in buildings){var split=r.Copy();split.Building=building;current.Details.Add(Row(split,metric,raw,1.0/buildings.Count,false,"Переход разделён поровну между корпусами",unique));}
                        }
                        else current.Details.Add(Row(r,metric,raw,item.Item3,false,item.Item5,unique));
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("AREA_BOOLEAN","Ошибка",ex.Message,r,metric.ToString());}
                }
            }
            private void CalculateCounts(List<Record> records,Indicator metric)
            {
                bool apartments=metric==Indicator.ApartmentsCount;
                string mode=apartments?Config.ApartmentMode:Config.ParkingMode;
                if(mode=="id"&&string.IsNullOrWhiteSpace(Config.Parameter(apartments?"apartment":"parking")))
                    throw new InvalidOperationException("Для подсчёта по ID задайте параметр идентификатора "+(apartments?"квартиры":"машино-места")+".");
                var keys=new HashSet<string>();
                foreach(var r in records.Where(x=>x.Source.Mode!="exclude").OrderBy(r=>r.Z).ThenBy(r=>r.Key,StringComparer.Ordinal))
                {
                    if(apartments&&!(r.Element is Room)||!apartments&&!IsParkingFamily(r.Element))continue;
                    if(r.Override==false||Eq(Mapped(r.Element,"include"),"0")||Eq(Mapped(r.Element,"include"),"нет"))continue;
                    if(apartments?(mode=="instances"?!(r.Element is FamilyInstance)||(r.Role!="apartment-family"&&r.Override!=true):string.IsNullOrWhiteSpace(r.Apartment)):r.Role!="parking")continue;
                    if(!apartments&&mode=="instances"&&!(r.Element is FamilyInstance))continue;
                    try
                    {
                        if(!apartments&&Above(r))continue;
                        if(r.Level==null||!r.Level.Include)continue;
                        string id=apartments?r.Apartment:Mapped(r.Element,"parking");
                        if(mode=="id"&&string.IsNullOrWhiteSpace(id))throw new InvalidOperationException("Не заполнен идентификатор "+(apartments?"квартиры":"машино-места")+".");
                        if(apartments&&string.IsNullOrWhiteSpace(r.Section))Notice("APARTMENT_SECTION","Предупреждение","Секция не задана. Номер квартиры должен быть уникален во всём корпусе.",r,metric.ToString());
                        string key=apartments?ApartmentKey(r):r.Source.Key+"|"+r.Building+"|"+r.Section+"|"+(mode=="id"?id:r.Element.UniqueId);
                        bool duplicate=!keys.Add(key);bool excluded=duplicate||r.Source.Mode=="reference";
                        if(duplicate&&!apartments)Notice("PARKING_DUPLICATE","Предупреждение","Повторный ID машино-места учтён один раз: "+id,r,metric.ToString());
                        Solid shape=null;try{if(ViewsEnabled)shape=r.Element is Room?RoomAreaBoundary(r):Plan(r,Indicator.PublicRooms);}catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("COUNT_GRAPHICS","Предупреждение","Количество определено, но нет контура: "+ex.Message,r,metric.ToString());}
                        current.Details.Add(Row(r,metric,1,1,excluded,duplicate?"Повторный объект той же квартиры / места":"Уникальный идентификатор: "+id,shape));
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("COUNT_INPUT","Ошибка",ex.Message,r,metric.ToString());}
                }
            }
            private void CalculateFloors(List<Record> records,Indicator metric)
            {
                foreach(var g in records.Where(r=>r.Source.Mode=="include"&&r.Level!=null&&(r.Element is SpatialElement||r.Element is Floor||OneOf(r.Role,"envelope","footprint","apartment-family"))).GroupBy(r=>r.Building+"|"+r.Section+"|"+Math.Round(r.Z,6)))
                {
                    try
                    {
                        var eligible=g.Where(r=>CountLevel(r,metric==Indicator.Storeys)).ToList();if(eligible.Count==0)continue;
                        current.Details.Add(Row(eligible[0],metric,1,1,false,"Один этаж на корпус и секцию; совпадающие отметки объединены"));
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("FLOOR_INPUT","Ошибка",ex.Message,g.First(),metric.ToString());}
                }
            }
            private void CalculateFootprints(List<Record> records)
            {
                foreach(var group in records.Where(r=>r.Source.Mode!="exclude").GroupBy(r=>new{r.Building,r.Section,Reference=r.Source.Mode=="reference",ReferenceKey=r.Source.Mode=="reference"?r.Source.Key:""}))
                {
                    var list=group.ToList();var metric=Indicator.Footprint;var inputs=new List<Tuple<Record,Solid>>();
                    try
                    {
                        var explicitGround=list.Where(r=>r.Role=="footprint"&&r.Override!=false).ToList();
                        if(explicitGround.Count>0)
                            foreach(var r in explicitGround){var plan=Plan(r,metric);if(plan!=null)inputs.Add(Tuple.Create(r,plan));}
                        else
                        {
                            var volume=BuildingVolume(list,metric.ToString());var lookup=list.ToDictionary(r=>r.Key,StringComparer.Ordinal);int processed=0;
                            foreach(var fragment in volume.Fragments)
                            {
                                var r=lookup[fragment.RecordKey];
                                Progress("Проекция застройки: "+(++processed)+" / "+volume.Fragments.Count+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                                Measure("Проекции и сечения для застройки",r,metric.ToString(),()=>
                                {
                                    var plan=Section(fragment.Shape,ground);if(plan!=null)inputs.Add(Tuple.Create(r,plan));
                                    var below=Half(fragment.Shape,ground,false);
                                    if(below!=null&&below.Volume>1e-9)inputs.Add(Tuple.Create(r,Projection(below)));
                                    return true;
                                });
                            }
                        }
                        foreach(var r in list.Where(r=>r.Override!=false))
                        {
                            bool residential=!IsPublic(r)||IsHigh(r);
                            bool extra=r.Role=="underground-footprint"||OneOf(r.Role,"porch","pit","passage","terrace","veranda","transition")||residential&&OneOf(r.Role,"balcony","loggia","entrance");
                            if(r.Role=="canopy")
                            {
                                var box=r.Element.get_BoundingBox(null);if(box==null)throw new InvalidOperationException("Не определена высота консоли над землёй.");
                                double lowest=Enumerable.Range(0,8).Select(i=>r.Source.Transform.OfPoint(box.Transform.OfPoint(new XYZ((i&1)==0?box.Min.X:box.Max.X,(i&2)==0?box.Min.Y:box.Max.Y,(i&4)==0?box.Min.Z:box.Max.Z))).Z).Min();
                                extra=lowest-ground<4.5/.3048;
                            }
                            if(extra){var plan=Plan(r,metric);if(plan!=null)inputs.Add(Tuple.Create(r,plan));}
                        }
                        var masks=new PlanIndex();foreach(var r in list.Where(r=>r.Override==false))masks.Add(Plan(r,metric),r);
                        if(inputs.Count==0)throw new InvalidOperationException("Пустой контур застройки на отметке земли. Проверьте отметку или задайте контур застройки.");
                        var used=new PlanIndex();int count=0;
                        foreach(var input in inputs.OrderBy(x=>x.Item1.Key,StringComparer.Ordinal))
                        {
                            var r=input.Item1;Progress("Фрагменты застройки: "+(++count)+" / "+inputs.Count+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                            var shape=RemoveNearby(input.Item2,masks,r,metric.ToString(),"Исключения площади застройки");
                            shape=RemoveNearby(shape,used,r,metric.ToString(),"Пересечения площади застройки");
                            if(shape==null||shape.Volume<1e-9)continue;
                            var row=Row(r,metric,shape.Volume*.09290304,1,group.Key.Reference,group.Key.Reference?"Справочный контур источника, без включения в итог":"Фрагмент контура застройки исходного объекта; пересечения учтены один раз",shape);
                            row.Level="План застройки";row.Elevation=ground;current.Details.Add(row);used.Add(shape,r);
                        }
                        if(explicitGround.Count>0&&!list.Any(r=>r.Role=="underground-footprint"))
                            Notice("FOOTPRINT_UNDERGROUND","Предупреждение","Задан ручной наземный контур без подземного. Подтвердите отсутствие выступающих подземных частей либо добавьте их контур.",list.First(),metric.ToString());
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("FOOTPRINT_INPUT","Ошибка",ex.Message,group.First(),metric.ToString());}
                }
            }
            private void ApplyDeltas(Indicator metric)
            {
                foreach(var c in Config.Corrections.Where(x=>x.Action=="delta"&&x.Metric==metric.ToString()))
                {
                    if(string.IsNullOrWhiteSpace(c.Reason))throw new InvalidOperationException("У корректирующего значения отсутствует обоснование.");
                    double delta=RequiredNumber(c.Value,"корректирующее значение");
                    if((int)metric>=21&&delta!=Math.Truncate(delta))throw new InvalidOperationException("Количество и этажность корректируются целыми числами.");
                    string building=string.IsNullOrWhiteSpace(c.Element)?"Ручная корректировка":c.Element,section=null;
                    var prior=current.Details.Where(d=>d.Metric==metric.ToString()&&!d.Excluded).ToList();
                    if(metric==Indicator.Storeys||metric==Indicator.Floors)
                    {
                        var target=(c.Element??"").Split('|');building=target[0];section=target.Length>1?target[1]:null;
                        var groups=prior.Where(d=>Eq(d.Building,building)&&(section==null||Eq(d.Section,section))).GroupBy(d=>d.Section).ToList();
                        if(groups.Count!=1)throw new InvalidOperationException("Для корректировки этажности укажите существующий корпус и секцию в формате Корпус|Секция.");
                        section=groups[0].Key;prior=groups[0].ToList();
                        if(prior.Sum(d=>d.Value)+delta<0)throw new InvalidOperationException("Корректировка приводит к отрицательному количеству этажей.");
                    }
                    c.Previous=prior.Sum(d=>d.Value).ToString("R",CultureInfo.InvariantCulture);
                    current.Details.Add(new Detail{Metric=metric.ToString(),Source=c.Source??"Ручной ввод",SourceKey=c.Source,Building=building,Section=section,
                        Level="Корректировка без контура",Raw=delta,Value=delta,Factor=1,Manual=true,Unit=Config.Metrics.First(x=>x.Key==metric.ToString()).Unit,Method=current.Method,
                        Reason=c.Reason+" | "+c.Author+" | "+c.Date,Profile=Config.Profile});
                    appliedCorrections.Add(c);
                }
            }
            private void Summarize(Metric m)
            {
                var rows=current.Details.Where(d=>d.Metric==m.Key&&!d.Excluded).ToList();
                bool floors=m.Key=="Storeys"||m.Key=="Floors";
                double value=floors?rows.GroupBy(d=>d.Building+"|"+d.Section).Select(g=>g.Sum(d=>d.Value)).DefaultIfEmpty(0).Max():rows.Sum(d=>d.Value);
                bool failed=rows.Count==0&&MetricBlocked(m.Key,current.Issues);
                current.Summary.Add(new Summary{Key=m.Key,Name=m.Name,Unit=m.Unit,Value=value,Method=current.Method,NotCalculated=failed,
                    Comment=floors?"Итог проекта - максимальное значение по корпусам / секциям. Поэтажные строки показывают состав.":"",Status=failed?"Не рассчитано":rows.Count==0?"Нет данных":"Рассчитано"});
                if(rows.Count==0)current.Issue("NO_DATA","Предупреждение","Нет включённых объектов показателя; ноль не подтверждает отсутствие таких объектов в проекте.",m.Key);
                if(value<0)current.Issue("NEGATIVE_TOTAL","Ошибка","Отрицательный итог после корректировок. Проверьте знак и величину дельты.",m.Key);
            }
        }
    }
}
