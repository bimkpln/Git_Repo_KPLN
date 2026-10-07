using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class StairFamilySetting : CategoryFamily
        {
            public bool WholeStair {get;set;}
            public string WidthParameter {get;set;} = "";
            public string WidthAxis {get;set;} = "";
        }

        public partial class Engine
        {
            private List<Record> automaticGeometryRecords = new List<Record>();
            private readonly Dictionary<string,double> automaticDimensions = new Dictionary<string,double>();
            private readonly Dictionary<string,string> automaticDimensionErrors = new Dictionary<string,string>();

            private static Level ModelLevel(Element element)
            {
                var room=element as Room;
                if(room!=null)return room.Level;
                var level=element.Document.GetElement(element.LevelId) as Level;
                if(level!=null)return level;
                var candidates=new List<Level>();
                foreach(var key in new[]{BuiltInParameter.FAMILY_LEVEL_PARAM,BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,BuiltInParameter.STAIRS_BASE_LEVEL_PARAM})
                {
                    var p=element.get_Parameter(key);
                    if(p?.StorageType!=StorageType.ElementId)continue;
                    level=element.Document.GetElement(p.AsElementId()) as Level;
                    if(level!=null&&!candidates.Any(l=>l.Id==level.Id))candidates.Add(level);
                }
                // No nearest-level inference: a hosted object may span several floors.
                return candidates.Count==1?candidates[0]:null;
            }

            private double AutomaticDimension(Record record,string field)
            {
                string key=record.Key+"/"+record.Role+"/"+record.Building+"/"+record.Section+"/"+field;
                double value;string failure;
                if(automaticDimensionErrors.TryGetValue(key,out failure))throw new InvalidOperationException(failure);
                if(!automaticDimensions.TryGetValue(key,out value))
                {
                    try
                    {
                        switch(field)
                        {
                            case "height": var clearance=RoomClearance(record);value=UniformHeight(clearance.Minimum,clearance.Maximum);break;
                            case "width": value=OpeningWidth(record)*.3048;break;
                            case "stair-width": value=AdjacentStairWidth(record);break;
                            case "roof-ratio": value=ObjectRoofRatio(record);break;
                            case "partial-floor": value=TechnicalPartialFloor(record)?1:0;break;
                            case "mezzanine-ratio": value=MezzanineRatio(record);break;
                            default: throw new InvalidOperationException("Неизвестный автоматический размер: "+field);
                        }
                        if(double.IsNaN(value)||double.IsInfinity(value)||value<0)throw new InvalidOperationException("Недопустимый результат измерения.");
                        automaticDimensions[key]=value;
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex)
                    {
                        string title=Config.Parameters.First(p=>p.Key==field).Title;
                        failure=title+": "+ex.Message;automaticDimensionErrors[key]=failure;
                        throw new InvalidOperationException(failure,ex);
                    }
                }
                string label=Config.Parameters.First(p=>p.Key==field).Title;
                string text=field=="partial-floor"?(value>0?"да":"нет"):value.ToString("0.######")+(field.Contains("width")||field=="height"?" м":"");
                bool manual=field=="roof-ratio"&&record.Level?.RoofRatioMode=="manual";
                if(manual)record.Manual=true;
                Notice("AUTO_DIMENSION",manual?"Предупреждение":"Информация",label+": "+text+(manual?". Ручное уточнение из таблицы этажей.":". Определено по модели и расчётным контурам."),record);
                return value;
            }

            // Only a constant-width rectangle is accepted. Neither an axis-aligned box nor a
            // convex hull of an irregular opening is evidence of its clear width.
            internal static double RectangularWidth(PlanarRegion region)
            {
                const double gridTolerance=4/PlanarRegion.Scale;
                if(region.Paths.Count!=1)throw new InvalidOperationException("Нужен один контур без отверстий. Проверьте расчётную область объекта.");
                var points=region.Paths[0].Select(p=>new[]{p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale}).ToList();
                bool removed=true;
                while(removed&&points.Count>4)
                {
                    removed=false;
                    for(int i=0;i<points.Count;i++)
                    {
                        var a=points[(i+points.Count-1)%points.Count];var b=points[i];var c=points[(i+1)%points.Count];
                        double x=b[0]-a[0],y=b[1]-a[1],u=c[0]-b[0],v=c[1]-b[1];
                        if(x*u+y*v>=0&&Math.Abs(x*v-y*u)<=gridTolerance*Math.Sqrt(x*x+y*y))
                        {points.RemoveAt(i);removed=true;break;}
                    }
                }
                if(points.Count!=4)throw new InvalidOperationException("У контура переменная или неоднозначная ширина. Автоматическое измерение поддерживает прямоугольный просвет; габарит вместо ширины не используется.");
                var lengths=new List<double>();
                for(int i=0;i<4;i++)
                {
                    var a=points[i];var b=points[(i+1)%4];var c=points[(i+2)%4];
                    double x=b[0]-a[0],y=b[1]-a[1],u=c[0]-b[0],v=c[1]-b[1];
                    double length=Math.Sqrt(x*x+y*y),next=Math.Sqrt(u*u+v*v);
                    if(length<=gridTolerance||Math.Abs(x*u+y*v)>gridTolerance*Math.Max(length,next))
                        throw new InvalidOperationException("Границы не образуют прямоугольный просвет постоянной ширины.");
                    lengths.Add(length);
                }
                return lengths.Min();
            }

            internal static double RectangularSpan(PlanarRegion region,double x,double y)
            {
                RectangularWidth(region);
                double length=Math.Sqrt(x*x+y*y);if(length<1e-9)throw new InvalidOperationException("Не задано направление ширины.");
                x/=length;y/=length;
                var path=region.Paths[0];var origin=path[0];
                for(int i=0;i<path.Count;i++)
                {
                    double dx=(path[(i+1)%path.Count].X-path[i].X)/PlanarRegion.Scale,dy=(path[(i+1)%path.Count].Y-path[i].Y)/PlanarRegion.Scale;
                    if(Math.Min(Math.Abs(dx*x+dy*y),Math.Abs(dx*y-dy*x))>4/PlanarRegion.Scale)
                        throw new InvalidOperationException("Направление ширины не совпадает со сторонами контура.");
                }
                var values=path.Select(p=>((p.X-origin.X)*x+(p.Y-origin.Y)*y)/PlanarRegion.Scale).ToList();
                return values.Max()-values.Min();
            }

            private double OpeningWidth(Record record)
            {
                var region=RoomNetRegion(record);
                if(record.Role=="stair-gap")return RectangularWidth(region);
                // A shallow arch can be wider than its depth. The shorter side is NOT its span.
                var family=record.Element as FamilyInstance;
                var wall=family?.Host as Wall;var axis=(wall?.Location as LocationCurve)?.Curve as Line;
                if(axis!=null)
                {
                    var direction=record.Source.Transform.OfVector(axis.Direction);
                    return RectangularSpan(region,direction.X,direction.Y);
                }
                var room=record.Element as Room;
                if(room!=null)
                {
                    using(var options=new SpatialElementBoundaryOptions{SpatialElementBoundaryLocation=SpatialElementBoundaryLocation.Finish})
                    {
                        var loops=room.GetBoundarySegments(options);
                        var sides=loops?.SelectMany(l=>l).Where(s=>room.Document.GetElement(s.ElementId) is Wall)
                            .Select(s=>s.GetCurve().CreateTransformed(record.Source.Transform)).OfType<Line>().ToList();
                        if(sides?.Count==2&&Math.Abs(Math.Abs(sides[0].Direction.DotProduct(sides[1].Direction))-1)<1e-9)
                        {
                            var normal=new XYZ(-sides[0].Direction.Y,sides[0].Direction.X,0);
                            double distance=Math.Abs((sides[1].GetEndPoint(0)-sides[0].GetEndPoint(0)).DotProduct(normal));
                            double span=RectangularSpan(region,normal.X,normal.Y);
                            if(Math.Abs(span-distance)<1e-6)return span;
                        }
                    }
                }
                throw new InvalidOperationException("Не определено направление ширины арочного проёма по его стенам. Глубина проёма не подставляется вместо ширины.");
            }

            private List<Record> GeometryPeers(Record record)
            {
                return automaticGeometryRecords.Where(r=>IsAreaInput(r.Element)&&r.Level!=null&&r.Level.Include&&r.Level.Kind!="exclude"&&r.Source.Mode=="include"&&r.Building==record.Building&&r.Section==record.Section).ToList();
            }

            private PlanarRegion ContainingFloor(Record floor,PlanarRegion target)
            {
                var components=RoomFloorRegions(automaticGeometryRecords,floor,Indicator.Gross)
                    .Where(p=>RegionIntersection(p,target).Area*.09290304>.001).ToList();
                if(components.Count!=1||RegionDifference(target,components[0]).Area*.09290304>.001)
                    throw new InvalidOperationException("Не определён единственный опорный контур этажа, охватывающий объект. Проверьте границы этажа и объекта.");
                return components[0];
            }

            private double ObjectRoofRatio(Record record)
            {
                if(record.Level==null)throw new InvalidOperationException("Не определён уровень надстройки.");
                if(record.Level.RoofRatioMode=="manual")
                {
                    if(automaticGeometryRecords.Where(r=>r.Level==record.Level&&IsAreaInput(r.Element)).Select(r=>r.Building+"/"+r.Section).Distinct().Count()>1)
                        throw new InvalidOperationException("Ручная доля уровня неоднозначна: на нём несколько корпусов или секций. Используйте геометрический расчёт.");
                    return ResolvedLevelField(record.Level,"ratio");
                }
                var scope=GeometryPeers(record).Where(r=>Math.Abs(r.Z-record.Z)<1e-6&&
                    (record.Level.Kind!="normal"||OneOf(r.Role,"roof-exit","roof-vent"))).ToList();
                if(scope.Count==0)throw new InvalidOperationException("Не определён состав надстройки в корпусе и секции объекта.");
                var hosts=Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude")
                    .SelectMany(s=>s.Elements.OfType<HostObject>().Where(e=>(e is Floor||e is RoofBase)&&PhaseAccepted(s,e)).Select(e=>Tuple.Create(s,e))).ToList();
                var review=new LevelGeometryReview();PrepareRoofRatio(record.Level,scope,hosts,review);
                Notice("AUTO_REFERENCE","Информация","Технические помещения надстройки: "+review.TechnicalArea?.ToString("0.###")+" м²; выбранная кровля: "+review.RoofArea?.ToString("0.###")+" м². "+string.Join("; ",review.Details),record);
                return review.Ratio??throw new InvalidOperationException("Не рассчитана доля кровли.");
            }

            private bool TechnicalPartialFloor(Record record)
            {
                var target=RoomNetRegion(record);var floor=ContainingFloor(record,target);
                var peers=GeometryPeers(record).Where(r=>Math.Abs(r.Z-record.Z)<1e-6&&RegionIntersection(RoomNetRegion(r),floor).Area>1e-9).ToList();
                var other=peers.FirstOrDefault(r=>OneOf(r.Role,"heated","auxiliary","public","parking")&&RegionIntersection(RoomNetRegion(r),floor).Area*.09290304>.001);
                if(other!=null)
                {
                    Notice("AUTO_REFERENCE","Информация","Техническое пространство занимает часть этажа: в том же контуре есть помещение другой функции, ID "+IDHelper.ElIdValue(other.Element.Id)+" ("+other.Source.Name+").",record);
                    return true;
                }
                if(peers.Count==0||peers.Any(r=>r.Role=="unknown"))throw new InvalidOperationException("Сначала распределите все помещения опорного этажа.");
                if(invalidRoomFloors.Contains(Math.Round(record.Z,6)))throw new InvalidOperationException("На опорном этаже есть пропущенные помещения с ошибками. Нельзя подтвердить, что техническое пространство занимает весь этаж.");
                var cells=PartitionRoomFloor(floor,peers);double technical=0;
                foreach(var cell in cells)
                {
                    var flags=cell.Owners.Select(r=>OneOf(r.Role,"technical-space","technical-void","technical-room")).Distinct().ToList();
                    if(flags.Count!=1)throw new InvalidOperationException("Общий участок этажа нельзя однозначно отнести к техническому пространству. Уточните расчётные области.");
                    if(flags[0])technical+=cell.Region.Area;
                }
                if(Math.Abs(cells.Sum(c=>c.Region.Area)-floor.Area)*.09290304>.001)
                    throw new InvalidOperationException("Классифицированные области не покрывают весь опорный этаж.");
                Notice("AUTO_REFERENCE","Информация","Техническое пространство: "+(technical*.09290304).ToString("0.###")+" м²; опорный этаж «"+record.Level.Name+"»: "+(floor.Area*.09290304).ToString("0.###")+" м².",record);
                return (floor.Area-technical)*.09290304>.001;
            }

            private double MezzanineRatio(Record record)
            {
                double z=RoomFloorElevation(record);var own=RoomNetRegion(record);var peers=GeometryPeers(record);
                // An underlying room must actually reach the mezzanine, not merely have the closest LevelId.
                var enclosing=new List<Record>();var sections=new Dictionary<Record,PlanarRegion>();var failures=new List<string>();
                foreach(var lower in peers.Where(r=>r.Role!="mezzanine"&&r.Element is Room&&r.Z<record.Z-1e-6))
                {
                    var box=ElementBounds(lower.Element,lower.Source.Transform);
                    if(box==null||box[5]<z-1e-6||RegionIntersection(BoundsRegion(box),own).IsEmpty)continue;
                    try{var section=SectionRegion(SpatialVolume(lower),z);if(!RegionIntersection(section,own).IsEmpty){enclosing.Add(lower);sections[lower]=section;}}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){failures.Add("ID "+IDHelper.ElIdValue(lower.Element.Id)+": "+ex.Message);}
                }
                var levels=enclosing.GroupBy(r=>Math.Round(r.Z,6)).ToList();
                if(levels.Count!=1||failures.Count>0)throw new InvalidOperationException("Не подтверждён единственный основной этаж антресоли по объёму нижних помещений. "+string.Join(" ",failures));
                if(RegionDifference(own,sections.Values.Aggregate(PlanarRegion.Empty,RegionUnion)).Area*.09290304>.001)
                    throw new InvalidOperationException("Нижние помещения не подтверждают опорный этаж для всей антресоли.");
                var baseRoom=levels[0].First();var floor=ContainingFloor(baseRoom,own);
                var mezzanines=peers.Where(r=>r.Role=="mezzanine"&&Math.Abs(RoomFloorElevation(r)-z)<1e-6&&RegionIntersection(RoomNetRegion(r),floor).Area>1e-9).ToList();
                var occupied=mezzanines.Select(RoomNetRegion).Aggregate(PlanarRegion.Empty,RegionUnion);
                var slab=PlanarRegion.Empty;
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"))
                    foreach(var element in source.Elements.OfType<Floor>())
                    {
                        if(!PhaseAccepted(source,element))continue;
                        var box=ElementBounds(element,source.Transform);
                        if(box==null||box[2]>z+1e-6||box[5]<z-1e-6||RegionIntersection(BoundsRegion(box),occupied).IsEmpty)continue;
                        foreach(var solid in Solids(element))foreach(Face face in solid.Faces)
                        {
                            var plane=face as PlanarFace;if(plane==null)continue;
                            if(source.Transform.OfVector(plane.FaceNormal).Z<1-1e-10||Math.Abs(source.Transform.OfPoint(plane.Origin).Z-z)>1e-6)continue;
                            var region=RegionFromLoops(plane.GetEdgesAsCurveLoops().Select(l=>CurveLoop.Create(l.Select(c=>c.CreateTransformed(source.Transform)).ToList())));
                            foreach(var component in region.Components().Select(c=>new PlanarRegion(c)))
                                if(!RegionIntersection(component,occupied).IsEmpty)slab=RegionUnion(slab,RegionIntersection(component,floor));
                        }
                    }
                if(slab.IsEmpty||RegionDifference(occupied,slab).Area*.09290304>.001)
                    throw new InvalidOperationException("Не найдена полная горизонтальная верхняя поверхность перекрытия антресоли на отметке её чистого пола.");
                if(peers.Any(r=>r.Role!="mezzanine"&&Math.Abs(r.Z-record.Z)<1e-6&&RegionIntersection(RoomNetRegion(r),slab).Area*.09290304>.001))
                    throw new InvalidOperationException("Перекрытие антресоли включает помещения другой категории; контур антресоли неоднозначен.");
                slab=ReviewRegion("normative/mezzanine/"+baseRoom.Level.Key+"/"+LevelGeometryScopeKey(mezzanines.Select(r=>r.Key)),record,"Антресоль - площадь перекрытия",z,()=>slab);
                if(RegionDifference(slab,floor).Area*.09290304>.001)throw new InvalidOperationException("Антресоль выходит за опорный контур этажа.");
                Notice("AUTO_REFERENCE","Информация","Антресоль: "+(slab.Area*.09290304).ToString("0.###")+" м²; основной этаж «"+baseRoom.Level.Name+"»: "+(floor.Area*.09290304).ToString("0.###")+" м².",record);
                return ValidRoofRatio(slab.Area,floor.Area);
            }

            private StairFamilySetting StairSetting(FamilyInstance instance)
            {return (Config.StairFamilies??new List<StairFamilySetting>()).FirstOrDefault(s=>Eq(s.Category,instance.Category?.Name)&&Eq(s.Family,instance.Symbol.FamilyName)&&Eq(s.Type,instance.Symbol.Name));}

            private PlanarRegion StairProjection(Element element,Source source)
            {
                return Solids(element).Select(s=>ProjectionRegion(SolidUtils.CreateTransformed(s,source.Transform))).Aggregate(PlanarRegion.Empty,RegionUnion);
            }

            private double FamilyRunWidth(FamilyInstance instance,StairFamilySetting setting,Source source)
            {
                if(!string.IsNullOrWhiteSpace(setting.WidthParameter))
                {
                    var p=Parameter(instance,setting.WidthParameter);double number;
                    if(p!=null&&p.HasValue&&p.StorageType==StorageType.Double&&IDHelper.IsLength(p))return p.AsDouble()*.3048;
                    if(p!=null&&p.HasValue&&p.StorageType==StorageType.String&&TryNumber(p.AsString(),out number))return number;
                    throw new InvalidOperationException("В семействе ID "+IDHelper.ElIdValue(instance.Id)+" не заполнена ширина марша «"+setting.WidthParameter+"» (Длина либо текст в метрах).");
                }
                if(setting.WholeStair)throw new InvalidOperationException("Для целой лестницы ID "+IDHelper.ElIdValue(instance.Id)+" выберите тегами типы вложенных маршей либо укажите параметр ширины марша.");
                if(!OneOf(setting.WidthAxis,"X","Y"))throw new InvalidOperationException("Для типа «"+setting.Caption+"» выберите направление ширины в локальных осях семейства либо параметр ширины марша.");
                var basis=setting.WidthAxis=="X"?instance.GetTransform().BasisX:instance.GetTransform().BasisY;
                var direction=source.Transform.OfVector(basis);
                if(Math.Abs(direction.Z)>1e-9)throw new InvalidOperationException("Направление ширины семейства не горизонтально.");
                return RectangularSpan(StairProjection(instance,source),direction.X,direction.Y)*.3048;
            }

            private double AdjacentStairWidth(Record record)
            {
                var gap=RoomNetRegion(record);double z=RoomFloorElevation(record);var widths=new List<double>();var errors=new List<string>();
                Action<Source,Element,Func<double>> measure=(source,element,read)=>
                {
                    var box=ElementBounds(element,source.Transform);
                    // One millimetre is only a contact tolerance, not a nearest-stair search radius.
                    const double contact=.001/.3048;
                    if(box==null||box[2]>z+contact||box[5]<z-contact)return;
                    var expanded=new[]{box[0]-contact,box[1]-contact,box[2],box[3]+contact,box[4]+contact,box[5]};
                    if(RegionIntersection(BoundsRegion(expanded),gap).IsEmpty)return;
                    try
                    {
                        var region=StairProjection(element,source);
                        if(!RegionsTouch(region,gap,contact))return;
                        double width=read();if(width<=0||double.IsNaN(width)||double.IsInfinity(width))throw new InvalidOperationException("Неположительная ширина марша.");
                        widths.Add(width);
                        Notice("AUTO_STAIR_SOURCE","Информация","Марш: "+source.Name+" / ID "+IDHelper.ElIdValue(element.Id)+"; ширина "+width.ToString("0.###")+" м.",record);
                    }
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){errors.Add("ID "+IDHelper.ElIdValue(element.Id)+": "+ex.Message);}
                };
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"))
                {
                    var visited=new HashSet<ElementId>();
                    foreach(var stair in source.Elements.OfType<Stairs>().Where(e=>PhaseAccepted(source,e)))
                        foreach(var id in stair.GetStairsRuns())
                        {
                            var run=source.Document.GetElement(id) as StairsRun;
                            if(run!=null&&visited.Add(id))measure(source,run,()=>run.ActualRunWidth*.3048);
                        }
                    Action<FamilyInstance> family=null;
                    family=instance=>
                    {
                        if(!visited.Add(instance.Id))return;
                        var setting=StairSetting(instance);if(setting==null)return;
                        var nested=instance.GetSubComponentIds().Select(id=>source.Document.GetElement(id)).OfType<FamilyInstance>().Where(c=>StairSetting(c)!=null).ToList();
                        if(setting.WholeStair&&string.IsNullOrWhiteSpace(setting.WidthParameter)&&nested.Count>0)
                        {foreach(var child in nested)family(child);return;}
                        measure(source,instance,()=>FamilyRunWidth(instance,setting,source));
                    };
                    foreach(var instance in source.Elements.OfType<FamilyInstance>().Where(e=>PhaseAccepted(source,e)))family(instance);
                }
                if(errors.Count>0)throw new InvalidOperationException(string.Join(" ",errors.Distinct()));
                if(widths.Count==0)throw new InvalidOperationException("Не найден примыкающий марш на отметке просвета. Проверьте теги в «Лестницы и марши» и геометрию просвета.");
                if(widths.Max()-widths.Min()>1e-6)throw new InvalidOperationException("К просвету примыкают марши разной ширины: "+string.Join("; ",widths.Distinct().Select(w=>w.ToString("0.###")))+" м. Единая ширина не подставляется.");
                return widths[0];
            }

            internal static bool RegionsTouch(PlanarRegion a,PlanarRegion b,double tolerance)
            {
                if(!RegionIntersection(a,b).IsEmpty)return true;
                Func<double[],double[],double[],double> distance=(p,x,y)=>
                {
                    double dx=y[0]-x[0],dy=y[1]-x[1],d=dx*dx+dy*dy;
                    double t=d==0?0:Math.Max(0,Math.Min(1,((p[0]-x[0])*dx+(p[1]-x[1])*dy)/d));
                    double u=p[0]-x[0]-t*dx,v=p[1]-x[1]-t*dy;return Math.Sqrt(u*u+v*v);
                };
                var aa=a.Paths.Select(p=>p.Select(q=>new[]{q.X/PlanarRegion.Scale,q.Y/PlanarRegion.Scale}).ToList()).ToList();
                var bb=b.Paths.Select(p=>p.Select(q=>new[]{q.X/PlanarRegion.Scale,q.Y/PlanarRegion.Scale}).ToList()).ToList();
                foreach(var x in aa)foreach(var y in bb)
                {
                    foreach(var p in x)for(int i=0;i<y.Count;i++)if(distance(p,y[i],y[(i+1)%y.Count])<=tolerance)return true;
                    foreach(var p in y)for(int i=0;i<x.Count;i++)if(distance(p,x[i],x[(i+1)%x.Count])<=tolerance)return true;
                }
                return false;
            }
        }
    }
}
