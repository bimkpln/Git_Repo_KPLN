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
            private sealed class ContourSelectionFilter : ISelectionFilter
            {
                internal bool Walls;
                public bool AllowElement(Element e){return Walls?e is Wall:e is FilledRegion&&(!Owned(e)||Kind(e)=="input-contour");}
                public bool AllowReference(Reference reference,XYZ point){return false;}
            }
            public void PerformContourAction()
            {
                string action=PendingContourAction;PendingContourAction=null;ReopenContours=true;
                try
                {
                    var metric=Config.Metrics.Single(m=>m.Key==ContourMetricKey);
                    if(action=="walls")
                    {
                        var picked=ui.Selection.PickObjects(ObjectType.Element,new ContourSelectionFilter{Walls=true},"Выберите стены активной модели и нажмите Готово");
                        metric.SelectedWalls=picked.Select(r=>doc.GetElement(r).UniqueId).Distinct().ToList();metric.WallSelectionMode="selection";metric.ContourMode="walls";return;
                    }
                    var regions=new List<FilledRegion>();
                    if(action.StartsWith("pick"))
                    {
                        var picked=ui.Selection.PickObjects(ObjectType.Element,new ContourSelectionFilter(),"Выберите цветовые области на поэтажных планах и нажмите Готово");
                        regions=picked.Select(r=>doc.GetElement(r)).Cast<FilledRegion>().Distinct().ToList();
                        foreach(var region in regions)
                            if(!((doc.GetElement(region.OwnerViewId) as ViewPlan)?.GenLevel is Level))throw new InvalidOperationException("Цветовая область должна находиться на плане с уровнем.");
                    }
                    else
                    {
                        var view=ui.ActiveView as ViewPlan;
                        if(view?.GenLevel==null||view.IsTemplate)throw new InvalidOperationException("Откройте поэтажный план нужного уровня перед запуском команды.");
                        var type=new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).FirstOrDefault(e=>!Owned(e))?.Id??ElementId.InvalidElementId;
                        if(type==ElementId.InvalidElementId)throw new InvalidOperationException("В проекте нет типа цветовой области. Создайте его средствами Revit.");
                        if(view.SketchPlane==null)
                            using(var t=new Transaction(doc,"ТЭП - плоскость ручного контура"))
                            {t.Start();view.SketchPlane=SketchPlane.Create(doc,Plane.CreateByNormalAndOrigin(XYZ.BasisZ,new XYZ(0,0,view.GenLevel.Elevation)));t.Commit();}
                        var points=new List<XYZ>();
                        while(true)
                        {
                            XYZ point;
                            try{point=ui.Selection.PickPoint(ObjectSnapTypes.Endpoints|ObjectSnapTypes.Intersections,"Вершины контура; Esc после третьей точки замыкает контур");}
                            catch(Autodesk.Revit.Exceptions.OperationCanceledException){if(points.Count<3)return;break;}
                            point=new XYZ(point.X,point.Y,view.GenLevel.Elevation);
                            if(points.Count>=3&&point.DistanceTo(points[0])<doc.Application.ShortCurveTolerance)break;
                            if(points.Count==0||point.DistanceTo(points.Last())>doc.Application.ShortCurveTolerance)points.Add(point);
                        }
                        var curves=points.Select((p,i)=>(Curve)Line.CreateBound(p,points[(i+1)%points.Count])).ToList();
                        var loops=new List<CurveLoop>{CurveLoop.Create(curves)};Extrude(loops);
                        using(var t=new Transaction(doc,"ТЭП - ручной расчётный контур"))
                        {t.Start();var region=FilledRegion.Create(doc,type,view.Id,loops);Tag(region,"input-contour",metric.Key);t.Commit();regions.Add(region);}
                    }
                    string kind=action.EndsWith("exclude")?"exclude":"base";
                    if(kind=="base"&&regions.Count>0)metric.ContourMode="manual";
                    foreach(var region in regions)
                    {
                        if(Config.Contours.Any(c=>c.Metric==metric.Key&&c.ElementUniqueId==region.UniqueId))continue;
                        var level=((ViewPlan)doc.GetElement(region.OwnerViewId)).GenLevel;
                        string profile=Config.Profile.StartsWith("by-")||Config.Profile=="high-mixed"?"":Config.Profile;
                        Config.Contours.Add(new ContourSketch{Metric=metric.Key,ElementUniqueId=region.UniqueId,ElementLabel=IDHelper.ElIdValue(region.Id).ToString(),
                            Kind=kind,LevelKey="host/"+level.UniqueId,LevelName=level.Name,
                            Building=Config.Grouping=="links"?doc.Title:Config.ManualBuilding,Profile=profile,
                            BuildingClass=profile.Contains("public")?"nonresidential":"residential"});
                    }
                }
                catch(Autodesk.Revit.Exceptions.OperationCanceledException){}
                catch(Exception ex){TaskDialog.Show("Контуры ТЭП",ex.Message);}
            }
            private static bool ContourSourceLevel(string key,string source)
            {return key!=null&&key.StartsWith(source+"/",StringComparison.Ordinal)&&key.IndexOf('/',source.Length+1)<0;}
            private LevelSetting ContourLevel(string key,string building,string section)
            {
                var settings=Config.Levels.Where(l=>l.Key==key).ToList();
                var specific=settings.Where(l=>(!string.IsNullOrWhiteSpace(l.Building)||!string.IsNullOrWhiteSpace(l.Section))&&
                    (string.IsNullOrWhiteSpace(l.Building)||Eq(l.Building,building))&&(string.IsNullOrWhiteSpace(l.Section)||Eq(l.Section,section))).ToList();
                if(specific.Count>1)throw new InvalidOperationException("Несколько настроек этажа соответствуют контуру: "+key);
                var result=specific.FirstOrDefault()??settings.FirstOrDefault(l=>string.IsNullOrWhiteSpace(l.Building)&&string.IsNullOrWhiteSpace(l.Section));
                if(result==null)throw new InvalidOperationException("Уровень контура отсутствует в настройках: "+key);return result;
            }
            private Record SketchRecord(ContourSketch sketch)
            {
                var region=doc.GetElement(sketch.ElementUniqueId) as FilledRegion;
                if(region==null)throw new InvalidOperationException("Цветовая область удалена или принадлежит другой модели: "+sketch.ElementLabel);
                var level=(doc.GetElement(region.OwnerViewId) as ViewPlan)?.GenLevel;
                if(level==null||sketch.LevelKey!="host/"+level.UniqueId)throw new InvalidOperationException("Изменился уровень ручного контура. Удалите привязку и выберите область заново.");
                if(string.IsNullOrWhiteSpace(sketch.Building)||!Choices("profile").Any(p=>p.Key==sketch.Profile&&!p.Key.StartsWith("by-")&&p.Key!="high-mixed"))
                    throw new InvalidOperationException("Укажите корпус и однозначный профиль ручного контура: "+sketch.ElementLabel);
                if(!OneOf(sketch.BuildingClass,"residential","nonresidential"))throw new InvalidOperationException("Укажите класс здания ручного контура.");
                return new Record{Element=region,Source=Sources.Single(s=>s.Key=="host"),Level=ContourLevel(sketch.LevelKey,sketch.Building,sketch.Section),
                    Building=sketch.Building,Section=sketch.Section??"",Profile=sketch.Profile,BuildingClass=sketch.BuildingClass,Role="heated",Part="auto",Manual=true};
            }
            private sealed class ContourInput
            {
                internal Record Record;internal Solid Plan;internal string Height;internal double Offset;internal bool Exclude;
            }
            private double ContourHeight(ContourInput input)
            {
                string value=string.IsNullOrWhiteSpace(input.Height)?input.Record.Level.Height:input.Height;
                if(!string.IsNullOrWhiteSpace(value))
                {double height=RequiredNumber(value,"высота участка, м")/.3048-(string.IsNullOrWhiteSpace(input.Height)?input.Offset:0);if(height<=1e-6)throw new InvalidOperationException("Высота участка должна быть положительной.");return height;}
                var r=input.Record;
                var next=Config.Levels.Where(l=>ContourSourceLevel(l.Key,r.Source.Key)&&l.Elevation>r.Z+1e-6)
                    .Select(l=>l.Key).Distinct().Select(k=>ContourLevel(k,r.Building,r.Section)).Where(l=>l.Include&&l.Kind!="exclude").OrderBy(l=>l.Elevation).FirstOrDefault();
                if(next==null)throw new InvalidOperationException("Нет следующего расчётного уровня. Укажите высоту верхнего участка на вкладке «Этажи» или у ручного контура.");
                double result=next.Elevation-r.Z-input.Offset;
                if(result<=1e-6)throw new InvalidOperationException("Смещение выреза находится выше следующего уровня. Задайте его высоту явно.");
                return result;
            }
            private static Solid ContourPrism(Solid plan,double bottom,double height)
            {
                // Disconnected islands are kept separate by the caller; holes remain in the face loops.
                return Extrude(BottomLoops(plan,bottom),height);
            }
            private static IEnumerable<Solid> PlanPieces(Solid plan)
            {
                foreach(var face in plan.Faces.Cast<Face>().OfType<PlanarFace>().Where(f=>f.FaceNormal.Z<-.999999))
                    yield return Extrude(face.GetEdgesAsCurveLoops().Select(l=>FlatLoop(l)).ToList());
            }
            private bool ContourWallSelected(Record r,Metric metric)
            {
                var wall=r.Element as Wall;if(wall==null)return false;
                if(metric.WallSelectionMode=="selection")return r.Source.Key=="host"&&metric.SelectedWalls.Contains(wall.UniqueId);
                if(metric.WallSelectionMode=="exterior")return wall.WallType.Function==WallFunction.Exterior;
                if(string.IsNullOrWhiteSpace(metric.WallParameter)||string.IsNullOrWhiteSpace(metric.WallValue))throw new InvalidOperationException("Задайте параметр и значение для отбора наружных стен.");
                string value=Value(wall,metric.WallParameter);return value!=null&&Eq(value,metric.WallValue);
            }
            private sealed class WallEdge
            {internal Record Record;internal Curve Curve;}
            private double WallOffset(WallEdge edge,Curve curve,Metric metric,Indicator indicator,double elevation)
            {
                var wall=(Wall)edge.Record.Element;
                string boundary=metric.WallBoundary;
                if(boundary=="auto")boundary=(int)indicator<=4?Methodology.GnsWallBoundary:GrossMetric(indicator)?"interior":"exterior";
                var exterior=WallFaceAtFloor(wall,edge.Record.Source.Transform,curve,elevation,ShellLayerType.Exterior);
                if(boundary=="exterior")return exterior.Offset;
                var interior=WallFaceAtFloor(wall,edge.Record.Source.Transform,curve,elevation,ShellLayerType.Interior);
                return boundary=="interior"?interior.Offset:WallCoreAtFloor(wall,exterior,interior,true);
            }
            private Solid AnalyticWallContour(List<Record> records,Metric metric,Indicator indicator,double? floorElevation=null)
            {
                double elevation=floorElevation??records.First().Z+RequiredNumber(metric.WallCutHeight,"высота сечения стен, м")/.3048;
                var remaining=records.Select(r=>new WallEdge{Record=r,Curve=Flat(((LocationCurve)r.Element.Location).Curve.CreateTransformed(r.Source.Transform))}).ToList();
                var loops=new List<CurveLoop>();const double tolerance=1e-5;
                // Degree two is required. Branches and gaps must be corrected rather than silently bridged.
                foreach(var edge in remaining)
                    for(int end=0;end<2;end++)
                    {
                        int count=remaining.Where(e=>!ReferenceEquals(e,edge)).Sum(e=>(e.Curve.GetEndPoint(0).DistanceTo(edge.Curve.GetEndPoint(end))<tolerance?1:0)+(e.Curve.GetEndPoint(1).DistanceTo(edge.Curve.GetEndPoint(end))<tolerance?1:0));
                        if(count!=1)throw new InvalidOperationException("Разрыв или ответвление наружных стен у ID "+IDHelper.ElIdValue(edge.Record.Element.Id)+". Нужен замкнутый контур без ответвлений.");
                    }
                while(remaining.Count>0)
                {
                    Progress("Сборка колец наружных стен: осталось "+remaining.Count);
                    var first=remaining[0];remaining.RemoveAt(0);
                    var curves=new List<Curve>{first.Curve};var offsets=new List<double>{WallOffset(first,first.Curve,metric,indicator,elevation)};
                    while(curves.Last().GetEndPoint(1).DistanceTo(curves[0].GetEndPoint(0))>=tolerance)
                    {
                        var end=curves.Last().GetEndPoint(1);
                        var edge=remaining.Single(e=>e.Curve.GetEndPoint(0).DistanceTo(end)<tolerance||e.Curve.GetEndPoint(1).DistanceTo(end)<tolerance);
                        var curve=edge.Curve.GetEndPoint(0).DistanceTo(end)<tolerance?edge.Curve:edge.Curve.CreateReversed();
                        curves.Add(curve);offsets.Add(WallOffset(edge,curve,metric,indicator,elevation));remaining.Remove(edge);
                    }
                    if(curves.Count<2)throw new InvalidOperationException("Недостаточно стен для замкнутого контура.");
                    var shifted=CurveLoop.CreateViaOffset(CurveLoop.Create(curves),offsets,XYZ.BasisZ);
                    if(shifted.Count()!=curves.Count)throw new InvalidOperationException("На отметке пола изменилось число участков наружного контура. Площадь не опубликована.");
                    loops.Add(shifted);
                }
                return Extrude(loops);
            }
            private List<ContourInput> ContourInputs(List<Record> records,Metric metric,Indicator indicator)
            {
                var result=new List<ContourInput>();
                if(metric.ContourMode=="walls")
                {
                    if(metric.WallSelectionMode=="selection")
                        foreach(var id in metric.SelectedWalls.Where(id=>!records.Any(r=>r.Source.Key=="host"&&r.Element.UniqueId==id)))
                            Notice("CONTOUR_WALL_MISSING","Ошибка","Выбранная стена отсутствует в доступном составе расчёта: "+id,metric:metric.Key);
                    var selected=records.Where(r=>ContourWallSelected(r,metric)).ToList();
                    double cut=RequiredNumber(metric.WallCutHeight,"высота сечения стен, м")/.3048;
                    if(cut<0)throw new InvalidOperationException("Высота сечения стен не может быть отрицательной.");
                    foreach(var group in selected.GroupBy(r=>new{r.Source.Key,r.Building,r.Section,r.Profile,r.BuildingClass}))
                    {
                        var seed=group.First();
                        var levels=Config.Levels.Where(l=>ContourSourceLevel(l.Key,seed.Source.Key)).Select(l=>l.Key).Distinct()
                            .Select(k=>ContourLevel(k,seed.Building,seed.Section)).Where(l=>l.Include&&l.Kind!="exclude").OrderBy(l=>l.Elevation);
                        foreach(var level in levels)
                        {
                            Progress("Наружный контур: "+seed.Building+" / "+level.Name);
                            var r=seed.Copy();r.Level=level;r.Override=null;r.Role="heated";r.Factor=null;
                            try
                            {
                                double z=level.Elevation+cut;
                                var onLevel=group.Where(w=>
                                {
                                    var box=w.Element.get_BoundingBox(null);if(box==null)return false;
                                    var a=w.Source.Transform.OfPoint(box.Transform.OfPoint(box.Min));var b=w.Source.Transform.OfPoint(box.Transform.OfPoint(box.Max));
                                    return z>=Math.Min(a.Z,b.Z)-1e-6&&z<Math.Max(a.Z,b.Z)-1e-6;
                                }).ToList();
                                if(onLevel.Count==0)
                                {
                                    if(records.Any(x=>x.Source.Key==r.Source.Key&&x.Building==r.Building&&x.Section==r.Section&&x.Level?.Key==level.Key&&x.Element is SpatialElement))
                                        Notice("CONTOUR_WALL_LEVEL","Ошибка","На этаже есть помещения, но на высоте сечения нет выбранных стен.",r,metric.Key);
                                    continue;
                                }
                                var plan=WallContour(onLevel,metric,indicator);
                                foreach(var piece in PlanPieces(plan))result.Add(new ContourInput{Record=r,Plan=piece});
                            }
                            catch(System.OperationCanceledException){throw;}
                            catch(Exception ex){Notice("CONTOUR_WALLS","Ошибка",level.Name+": "+ex.Message,r,metric.Key);}
                        }
                    }
                }
                foreach(var sketch in Config.Contours.Where(c=>c.Metric==metric.Key&&(c.Kind=="exclude"||metric.ContourMode=="manual")))
                {
                    Record r=null;
                    try
                    {
                        r=SketchRecord(sketch);if(r.Source.Mode=="exclude"||!r.Level.Include||r.Level.Kind=="exclude")continue;
                        double offset=string.IsNullOrWhiteSpace(sketch.BottomOffset)?0:RequiredNumber(sketch.BottomOffset,"смещение низа выреза, м")/.3048;
                        if(sketch.Kind=="base"&&Math.Abs(offset)>1e-8)throw new InvalidOperationException("Смещение низа допускается только у выреза. Основной контур начинается на отметке своего уровня.");
                        var shape=Extrude(((FilledRegion)r.Element).GetBoundaries().Select(l=>FlatLoop(l)).ToList());
                        foreach(var piece in PlanPieces(shape))result.Add(new ContourInput{Record=r,Plan=piece,Height=sketch.Height,Offset=offset,Exclude=sketch.Kind=="exclude"});
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("CONTOUR_SKETCH","Ошибка","Область "+sketch.ElementLabel+": "+ex.Message,r,metric.Key);if(sketch.Kind=="exclude")throw new InvalidOperationException("Ручной вырез недоступен. Расчёт по этому показателю остановлен до восстановления выреза.",ex);}
                }
                if(!result.Any(i=>!i.Exclude))Notice("CONTOUR_EMPTY","Ошибка","Не найдено ни одного основного контура. Проверьте выбор стен / областей, уровни и источники.",metric:metric.Key);
                return result;
            }
            private bool ContourMaskRequired(Record r,Metric metric,Indicator indicator,List<Record> all)
            {
                if(r.Override.HasValue)return !r.Override.Value;
                string include=Mapped(r.Element,"include");if(Eq(include,"0")||Eq(include,"нет")||Eq(include,"false"))return true;
                if(metric.ContourExclusions=="rules"||indicator==Indicator.Footprint)return false;
                if(OneOf(r.Role,"structure","envelope","footprint","underground-footprint"))return false;
                if(r.Role=="unknown")
                {if(r.Element is SpatialElement)throw new InvalidOperationException("Назначение помещения / зоны не определено. Нельзя проверить нормативные вырезы внутри контура.");return false;}
                if(metric.Key.StartsWith("Volume"))return OneOf(r.Role,"balcony","terrace","canopy","passage","decoration","ventilated-void","soil-filled");
                if(!Eligible(r,indicator))return r.Element is SpatialElement||OneOf(r.Role,"multilight","stair-gap","opening","shaft","engineering-shaft","stove","decoration");
                if(OneOf(r.Role,"multilight","stair-gap","opening","shaft","engineering-shaft"))return VerticalExclusion(r,indicator,all,ContourMaskPlan(r,indicator));
                return false;
            }
            private Solid ContourMaskPlan(Record r,Indicator indicator)
            {
                string key=r.Key+"/contour-mask";Solid plan;
                if(!shapes.TryGetValue(key,out plan))
                {
                    plan=r.Element is SpatialElement?SpatialPlan(r,"net",indicator.ToString()):Plan(r,Indicator.Footprint);
                    if(plan==null||plan.Volume<1e-9)throw new InvalidOperationException("Пустая геометрия исключаемого объекта.");
                    shapes[key]=plan;
                }
                return plan;
            }
            private bool ContourBaseIncluded(Record r,Indicator indicator)
            {
                if(r.Level==null||!r.Level.Include||r.Level.Kind=="exclude")return false;
                if(indicator==Indicator.Footprint||indicator.ToString().StartsWith("Volume"))return true;
                var probe=r.Copy();probe.Role=r.Level.Kind=="attic"?"attic":r.Level.Kind=="void"?"technical-void":r.Level.Kind=="roof"?"roof-used":r.Level.Kind=="technical"?"technical-room":"heated";
                // A wall's inclusion parameter must not discard an entire floor envelope.
                var key=r.Level.Key.Substring(r.Source.Key.Length+1);probe.Element=r.Source.Document.GetElement(key)??r.Element;
                probe.Override=null;return Eligible(probe,indicator);
            }
            private void CalculateContours(List<Record> records,Indicator indicator,Metric metric)
            {
                bool volume=metric.Key.StartsWith("Volume"),footprint=indicator==Indicator.Footprint;
                var inputs=ContourInputs(records,metric,indicator);var bases=inputs.Where(i=>!i.Exclude).ToList();
                if(volume)Notice("CONTOUR_PRISMS","Предупреждение","Расчёт по вертикальным поэтажным призмам: контур постоянен по высоте участка. Наклонные кровли, наклонные стены и изменения сечения внутри этажа этим режимом не восстанавливаются. Проверьте высоты и 3D либо используйте текущий способ по геометрии.",metric:metric.Key);
                if(metric.ContourExclusions=="rules")Notice("CONTOUR_EXPLICIT_ONLY","Предупреждение","Автоматические исключения по назначению отключены. Применяются только явные правила, параметр включения и ручные вырезы.",metric:metric.Key);
                if(metric.ContourExclusions=="normative"&&!records.Any(r=>r.Element is SpatialElement))
                    Notice("CONTOUR_NO_ROOMS","Предупреждение","В составе расчёта нет помещений, пространств или зон для проверки исключений. Проверьте полноту вырезов вручную.",metric:metric.Key);
                var maskRecords=new List<Record>();var badGroups=new HashSet<string>();
                Func<Record,string> groupKey=r=>r.Building+"|"+r.Section+"|"+(r.Source.Mode=="reference"?r.Source.Key:"include");
                foreach(var r in records.Where(r=>r.Level!=null&&r.Level.Include&&r.Level.Kind!="exclude"))
                {
                    if(!bases.Any(b=>groupKey(b.Record)==groupKey(r)&&(volume||footprint||Math.Abs(b.Record.Z-r.Z)<1e-6)))continue;
                    try{if(ContourMaskRequired(r,metric,indicator,records))maskRecords.Add(r);}
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("CONTOUR_CLASSIFICATION","Ошибка",ex.Message,r,metric.Key);badGroups.Add(groupKey(r));}
                }
                var used=new Dictionary<string,PlanIndex>();int count=0;
                foreach(var input in bases.OrderBy(i=>i.Record.Z).ThenBy(i=>i.Record.Key,StringComparer.Ordinal))
                {
                    var r=input.Record;string key=groupKey(r);
                    Progress("Контур и вырезы: "+metric.Name+" - "+(++count)+" / "+bases.Count+"; "+r.Level.Name);
                    try
                    {
                        if(!ContourBaseIncluded(r,indicator)||badGroups.Contains(key))continue;
                        if(footprint&&r.Z>ground+1e-6)continue;
                        double height=volume?ContourHeight(input):1;
                        Solid shape=volume?ContourPrism(input.Plan,r.Z,height):input.Plan;
                        // Footprint uses the union of contours at/below ground; upper protrusions must be supplied as explicit ground outlines.
                        var maskIndex=new PlanIndex(volume);var exclusions=new List<Detail>();
                        Action<Record,Solid,string> addMask=(owner,mask,reason)=>
                        {
                            if(mask==null||mask.Volume<1e-9)return;
                            maskIndex.Add(mask,owner);
                            var intersections=volume?VolumeIntersections(shape,mask):new List<Solid>{Intersect(shape,mask)};
                            foreach(var part in intersections)
                            {
                                var clipped=part;if(clipped==null||clipped.Volume<1e-9)continue;
                                if(volume&&indicator!=Indicator.Volume)clipped=Half(clipped,zero,indicator==Indicator.VolumeAbove);
                                if(clipped==null||clipped.Volume<1e-9)continue;
                                var row=Row(owner,indicator,clipped.Volume*(volume?.028316846592:.09290304),1,true,reason,clipped,volume);
                                row.Level=r.Level.Name;row.Elevation=footprint?ground:r.Z;exclusions.Add(row);
                            }
                        };
                        foreach(var mask in maskRecords.Where(m=>groupKey(m)==key&&(volume||footprint||Math.Abs(m.Z-r.Z)<1e-6)))
                        {
                            Progress("Вырезы: "+r.Level.Name+"; ID "+IDHelper.ElIdValue(mask.Element.Id));
                            if(volume&&metric.VolumeMaskHeight=="actual")
                            {
                                if(mask.Element is Area)throw new InvalidOperationException("Зона ID "+IDHelper.ElIdValue(mask.Element.Id)+" не имеет объёма. Выберите вырез на всю высоту участка или задайте ручной вырез.");
                                var solids=mask.Element is SpatialElement?new List<Solid>{SpatialVolume(mask)}:VolumeSolids(mask,metric.Key);
                                if(solids.Count==0)throw new InvalidOperationException("Нет объёмной геометрии выреза ID "+IDHelper.ElIdValue(mask.Element.Id));
                                foreach(var solid in solids)addMask(mask,solid,"Исключение по фактической геометрии объекта");
                            }
                            else
                            {
                                if(volume&&Math.Abs(mask.Z-r.Z)>1e-6)continue;
                                foreach(var plan in PlanPieces(ContourMaskPlan(mask,indicator)))
                                    addMask(mask,volume?ContourPrism(plan,r.Z,height):plan,volume?"Исключение на всю высоту участка":"Исключение по контуру помещения / элемента");
                            }
                        }
                        foreach(var mask in inputs.Where(m=>m.Exclude&&groupKey(m.Record)==key&&(volume||footprint||Math.Abs(m.Record.Z-r.Z)<1e-6)))
                        {
                            var solid=volume?ContourPrism(mask.Plan,mask.Record.Z+mask.Offset,ContourHeight(mask)):mask.Plan;
                            addMask(mask.Record,solid,"Ручной вырез цветовой областью");
                        }
                        string usedKey=key+(volume||footprint?"":"|"+r.Z.ToString("R",CultureInfo.InvariantCulture));
                        PlanIndex previous;if(!used.TryGetValue(usedKey,out previous))used[usedKey]=previous=new PlanIndex(volume);
                        var fragments=volume?RemoveNearbyVolumes(new List<Solid>{shape},maskIndex,r,metric.Key,"Исключения внутри контура"):new List<Solid>{RemoveNearby(shape,maskIndex,r,metric.Key,"Вырезы контура")};
                        var final=new List<Solid>();
                        foreach(var fragment in fragments)
                        {
                            if(fragment==null||fragment.Volume<1e-9)continue;
                            if(volume)final.AddRange(RemoveNearbyVolumes(new List<Solid>{fragment},previous,r,metric.Key,"Устранение повторного учёта контуров"));
                            else{var unique=RemoveNearby(fragment,previous,r,metric.Key,"Пересечения контуров");if(unique!=null&&unique.Volume>1e-9)final.Add(unique);}
                        }
                        // Publish the whole input atomically, including cuts at zero, before indexing it as counted.
                        var rows=new List<Detail>();
                        foreach(var fragment in final)
                        {
                            var measured=volume&&indicator!=Indicator.Volume?Half(fragment,zero,indicator==Indicator.VolumeAbove):fragment;
                            if(measured==null||measured.Volume<1e-9)continue;
                            var row=Row(r,indicator,measured.Volume*(volume?.028316846592:.09290304),1,r.Source.Mode=="reference",
                                (metric.ContourMode=="walls"?"Контур наружных стен":"Ручной контур")+"; вырезы и пересечения вычтены"+(volume?"; высота участка "+(height*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" м":""),measured,volume);
                            row.Purpose="envelope";
                            if(footprint){row.Level="План застройки";row.Elevation=ground;}rows.Add(row);
                        }
                        foreach(var fragment in final)previous.Add(fragment,r);
                        current.Details.AddRange(exclusions);current.Details.AddRange(rows);
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("CONTOUR_CALCULATION","Ошибка",r.Level.Name+": "+ex.Message,r,metric.Key);}
                }
                if(footprint&&!bases.Any(b=>b.Record.Source.Mode=="include"&&b.Record.Z<=ground+1e-6))
                    Notice("CONTOUR_GROUND_EMPTY","Ошибка","Нет основного контура на отметке земли или ниже. Проверьте отметку земли и уровни областей.",metric:metric.Key);
                if(footprint)Notice("CONTOUR_FOOTPRINT_EXTRAS","Предупреждение","В этом режиме площадь застройки определяется выбранными контурами: наземным сечением и проекциями подземных этажей. Крыльца, приямки, консоли и другие выступающие части добавьте в ручные основные области либо используйте текущий способ.",metric:metric.Key);
            }
        }
    }
}
