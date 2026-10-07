using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
    public class LevelRoofOption
    {
        public string Key {get;set;}
        public string Name {get;set;}
        public double Area {get;set;}
        public string Display {get{return Name+" | "+Area.ToString("0.###")+" м²";}}
        internal PlanarRegion Region;
        internal Record Record;
    }
    public class LevelGeometryReview
    {
        public double? HeightMin {get;set;}
        public double? HeightMax {get;set;}
        public double? AnnexArea {get;set;}
        public double? TechnicalArea {get;set;}
        public double? RoofArea {get;set;}
        public double? Ratio {get;set;}
        public List<LevelRoofOption> Roofs {get;set;}=new List<LevelRoofOption>();
        public Dictionary<string,string> Errors {get;set;}=new Dictionary<string,string>();
        public List<string> Details {get;set;}=new List<string>();
    }
    public partial class Engine
    {
        private readonly Dictionary<string,LevelGeometryReview> levelGeometryReviews=new Dictionary<string,LevelGeometryReview>();
        public LevelGeometryReview LevelGeometryFor(LevelSetting level)
        {LevelGeometryReview review;return levelGeometryReviews.TryGetValue(level.Key,out review)?review:new LevelGeometryReview();}
        public static double ValidRoofRatio(double numerator,double denominator)
        {
            if(double.IsNaN(numerator)||double.IsInfinity(numerator)||double.IsNaN(denominator)||double.IsInfinity(denominator)||numerator<0||denominator<=0||numerator>denominator)
                throw new InvalidOperationException("Площади не подтверждают долю кровли: проверьте числитель, знаменатель и выбранную кровлю.");
            return numerator/denominator;
        }
        public static double UniformHeight(double min,double max)
        {
            if(double.IsNaN(min)||double.IsNaN(max)||double.IsInfinity(min)||double.IsInfinity(max)||min<=0||max<min)
                throw new InvalidOperationException("Не определена положительная высота.");
            if(max-min>1e-6)throw new InvalidOperationException("Переменная высота "+min.ToString("0.###")+" - "+max.ToString("0.###")+" м. Уточните применяемое значение; среднее автоматически не используется.");
            return min;
        }
        private double ResolvedLevelField(LevelSetting level,string field)
        {
            string mode=field=="height"?level.HeightMode:field=="area"?level.RoofAreaMode:level.RoofRatioMode;
            if(mode=="manual")
            {
                double value=RequiredNumber(field=="height"?level.Height:field=="area"?level.RoofArea:level.RoofRatio,"ручное значение «"+field+"» этажа «"+level.Name+"»");
                if(value<0||field!="ratio"&&value==0||field=="ratio"&&value>1)throw new InvalidOperationException("Ручное значение вне допустимого диапазона.");
                return value;
            }
            var review=LevelGeometryFor(level);string error;
            if(review.Errors.TryGetValue(field,out error))throw new InvalidOperationException("Этаж «"+level.Name+"»: "+error);
            if(field=="height"&&review.HeightMin.HasValue&&review.HeightMax.HasValue)return UniformHeight(review.HeightMin.Value,review.HeightMax.Value);
            var result=field=="area"?review.AnnexArea:field=="ratio"?review.Ratio:null;
            if(!result.HasValue)throw new InvalidOperationException("Этаж «"+level.Name+"»: геометрия для «"+field+"» не определена. Откройте значение в таблице этажей.");
            return result.Value;
        }
        public void PreviewLevelGeometry(Action<string> progress=null)
        {
            var saved=current;var savedProgress=reportProgress;bool savedAutomatic=automaticAreaInputs;
            current=new Run();reportProgress=progress;automaticAreaInputs=false;
            try
            {
                phaseCache.Clear();parameterCache.Clear();shapes.Clear();DisposeSpatialCalculators();localSpatialVolumes.Clear();
                measuredRoomRegions.Clear();roomNetRegions.Clear();roomFloorRegions.Clear();roomFloorFailures.Clear();
                WithPreservedReviewInputs(()=>{BeginReviewInputs();RefreshLevelGeometry();});
            }
            finally{current=saved;reportProgress=savedProgress;automaticAreaInputs=savedAutomatic;DisposeSpatialCalculators();localSpatialVolumes.Clear();shapes.Clear();}
        }
        private void WithPreservedReviewInputs(Action preview)
        {
            // A settings preview must not replace the last calculation's navigable areas.
            var inputs=reviewInputs.ToList();
            var bindings=reviewBindings.ToList();
            var overlays=editedObjectOverlays.ToList();
            string configuration=reviewConfiguration;
            try{preview();}
            finally
            {
                reviewInputs.Clear();foreach(var input in inputs)reviewInputs.Add(input.Key,input.Value);
                reviewBindings=bindings;
                editedObjectOverlays.Clear();foreach(var overlay in overlays)editedObjectOverlays.Add(overlay.Key,overlay.Value);
                reviewConfiguration=configuration;
            }
        }
        private List<Record> LevelGeometryRecords(List<Source> sources,Dictionary<string,List<string>> errors)
        {
            var result=new List<Record>();
            foreach(var source in sources)foreach(var element in source.Elements)
            {
                var room=element as Room;
                if(room!=null&&(!IsPlacedRoom(room)||room.Area<=1e-9))continue;
                if(room==null&&!IsClassifiedFamily(element))continue;
                var level=ModelLevel(element);if(level==null)continue;
                var setting=Config.Levels.FirstOrDefault(l=>l.Key==source.Key+"/"+level.UniqueId&&l.Include);if(setting==null)continue;
                try{if(!PhaseAccepted(source,element))continue;}
                catch(OperationCanceledException){throw;}
                catch(Exception ex)
                {List<string> failures;if(!errors.TryGetValue(setting.Key,out failures))errors[setting.Key]=failures=new List<string>();failures.Add("ID "+IDHelper.ElIdValue(element.Id)+": "+ex.Message);continue;}
                string role;try{role=room!=null?DepartmentRole(ClassificationValue(room)):ClassifiedFamilyRole(element);}catch{role="unknown";}
                if(room!=null&&CategoryUsesFamilies(role))continue;
                string profile=Config.Profile;
                try{var mapped=Mapped(element,"profile");if(!string.IsNullOrWhiteSpace(mapped))profile=Choices("profile").FirstOrDefault(p=>Eq(p.Key,mapped)||Eq(p.Label,mapped))?.Key??profile;}catch{ }
                result.Add(new Record{Source=source,Element=element,Level=setting,Profile=profile,Role=role,Part="auto"});
            }
            return result;
        }
        private void RefreshLevelGeometry()
        {
            levelGeometryReviews.Clear();
            var sources=Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude").ToList();
            var inputErrors=new Dictionary<string,List<string>>();var records=LevelGeometryRecords(sources,inputErrors);
            var hosts=new List<Tuple<Source,HostObject>>();
            foreach(var source in sources)foreach(var host in source.Elements.OfType<HostObject>().Where(e=>e is Floor||e is RoofBase||e is Ceiling))
            {try{if(PhaseAccepted(source,host))hosts.Add(Tuple.Create(source,host));}catch(OperationCanceledException){throw;}catch(Exception){/* The affected source's rooms carry the stage error; no structural datum is inferred. */}}
            var hostCache=new Dictionary<string,List<HeightPatch>>();
            foreach(var level in Config.Levels)
            {
                var onLevel=records.Where(r=>r.Level==level).ToList();
                var profiles=onLevel.Select(r=>r.Profile).Distinct().ToList();
                level.NormativeProfile=profiles.Count==1?profiles[0]:Config.Profile;level.HasRoofObjects=onLevel.Any(r=>OneOf(r.Role,"roof-exit","roof-vent"));
                level.HasHighObjects=profiles.Any(p=>p.StartsWith("high-"));level.HasPublicObjects=profiles.Contains("public");
                if(!level.Include||!(level.HeightRequired||level.RoofAreaRequired||level.RoofRatioRequired))continue;
                Progress("Геометрия подполья / надстройки: "+level.Source+" / "+level.Name);
                var review=new LevelGeometryReview();levelGeometryReviews[level.Key]=review;
                var scope=level.Kind=="normal"?onLevel.Where(r=>OneOf(r.Role,"roof-exit","roof-vent")).ToList():onLevel;
                foreach(var r in scope)review.Details.Add(r.Source.Name+" / "+r.Element.Name+" / ID "+IDHelper.ElIdValue(r.Element.Id));
                Action<string,Action> attempt=(field,action)=>
                {try{if(scope.Count==0)throw new InvalidOperationException("Нет размещённых помещений или классифицированных семейств для этой проверки.");
                    List<string> failures;if(inputErrors.TryGetValue(level.Key,out failures))throw new InvalidOperationException(string.Join(" ",failures.Distinct()));
                    if(profiles.Count>1)throw new InvalidOperationException("На уровне разные профили здания. Общие высота и площади для нормативной проверки неоднозначны; уточните состав уровня.");action();}
                    catch(OperationCanceledException){throw;}catch(Exception ex){review.Errors[field]=ex.Message;}};
                if(level.HeightRequired)attempt("height",()=>
                {
                    var ranges=new List<double[]>();
                    foreach(var r in scope)
                    {
                        Progress("Высота: "+r.Source.Name+" / ID "+IDHelper.ElIdValue(r.Element.Id));
                        var patches=r.Element is Room?HeightPatches(SpatialVolume(r),Transform.Identity):
                            Solids(r.Element).SelectMany(s=>HeightPatches(s,r.Source.Transform)).ToList();
                        if(r.Element is Room)ConfirmRoomHeight(patches,r,hosts,hostCache);
                        ranges.Add(HeightRange(patches));
                    }
                    review.HeightMin=ranges.Min(r=>r[0]);review.HeightMax=ranges.Max(r=>r[1]);
                    UniformHeight(review.HeightMin.Value,review.HeightMax.Value);
                });
                if(level.RoofAreaRequired)attempt("area",()=>
                {
                    var first=scope[0];
                    var area=ReviewRegion("normative/annex/"+level.Key+"/"+LevelGeometryScopeKey(scope.Select(r=>r.Key)),first,"Надстройка - наружный контур",first.Z,()=>AnnexOutline(scope));
                    if(area.IsEmpty)throw new InvalidOperationException("Пустой контур надстройки.");
                    review.AnnexArea=area.Area*.09290304;
                });
                if(level.RoofRatioRequired)attempt("ratio",()=>PrepareRoofRatio(level,scope,hosts,review));
                level.HeightSummary=LevelFieldSummary(level,"height",review);
                level.RoofAreaSummary=LevelFieldSummary(level,"area",review);
                level.RoofRatioSummary=LevelFieldSummary(level,"ratio",review);
            }
        }
        private string LevelFieldSummary(LevelSetting level,string field,LevelGeometryReview review)
        {
            try
            {
                double value=ResolvedLevelField(level,field);
                bool manual=(field=="height"?level.HeightMode:field=="area"?level.RoofAreaMode:level.RoofRatioMode)=="manual";
                return (manual?"Вручную: ":"")+value.ToString(field=="ratio"?"0.##%":"0.###");
            }
            catch
            {
                if(field=="height"&&review.HeightMin.HasValue&&review.HeightMax.HasValue)
                    return review.HeightMin.Value.ToString("0.###")+" - "+review.HeightMax.Value.ToString("0.###");
                return "Задать";
            }
        }
        public List<InputAuditRow> LevelGeometryAuditRows()
        {
            var rows=new List<InputAuditRow>();
            foreach(var level in Config.Levels.Where(l=>l.Include))
            {
                var review=LevelGeometryFor(level);
                foreach(var field in new[]{"height","area","ratio"})
                {
                    if(!(field=="height"?level.HeightRequired:field=="area"?level.RoofAreaRequired:level.RoofRatioRequired))continue;
                    bool manual=(field=="height"?level.HeightMode:field=="area"?level.RoofAreaMode:level.RoofRatioMode)=="manual";
                    string message=string.Join("; ",review.Details);bool problem=manual;
                    try{ResolvedLevelField(level,field);}
                    catch(Exception ex){problem=true;message=ex.Message;}
                    if(field=="ratio")message="Технические помещения: "+(review.TechnicalArea?.ToString("0.###")??"не определено")+" м²; кровля: "+(review.RoofArea?.ToString("0.###")??"не определено")+" м². "+message;
                    if(manual)message="Применяется ручное значение. "+message;
                    rows.Add(new InputAuditRow{Source=level.Source,ElementId="-",Name=level.Name,Parameter=field=="height"?"Высота":field=="area"?"Площадь надстройки":"Доля кровли",
                        Value=LevelFieldSummary(level,field,review),State=message,Problem=problem});
                }
            }
            return rows;
        }
        private void ReportLevelGeometry()
        {
            foreach(var row in LevelGeometryAuditRows())current.Issues.Add(new Issue{Code="LEVEL_GEOMETRY",Source=row.Source,
                Severity=row.Problem?"Предупреждение":"Информация",Message=row.Name+" / "+row.Parameter+": "+row.Value+". "+row.State});
        }
        private PlanarRegion AnnexOutline(List<Record> scope)
        {
            var result=PlanarRegion.Empty;
            foreach(var family in scope.Where(r=>!(r.Element is Room)))result=RegionUnion(result,FamilyAreaRegion(family));
            var rooms=scope.Where(r=>r.Element is Room).ToList();if(rooms.Count==0)return result;
            double z=RoomFloorElevation(rooms[0]);
            if(rooms.Any(r=>Math.Abs(RoomFloorElevation(r)-z)>FloorGeometryTolerance))throw new InvalidOperationException("Разные отметки пола надстройки. Требуется проверенный общий контур.");
            var footprints=rooms.Select(RoomNetRegion).Aggregate(PlanarRegion.Empty,RegionUnion);
            var points=footprints.Paths.SelectMany(p=>p).ToList();double minX=points.Min(p=>p.X)/PlanarRegion.Scale,maxX=points.Max(p=>p.X)/PlanarRegion.Scale,minY=points.Min(p=>p.Y)/PlanarRegion.Scale,maxY=points.Max(p=>p.Y)/PlanarRegion.Scale;
            var walls=Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude").SelectMany(s=>s.Elements.OfType<Wall>().SelectMany(WallMembers).Distinct()
                .Where(w=>w.WallType.Function==WallFunction.Exterior&&PhaseAccepted(s,w)&&CrossesFloor(w,s.Transform,z))
                .Select(w=>new Record{Source=s,Element=w,Level=rooms[0].Level})).Where(r=>
                    {var b=ElementBounds(r.Element,r.Source.Transform);return b!=null&&b[3]>=minX-FloorGeometryTolerance&&b[0]<=maxX+FloorGeometryTolerance&&b[4]>=minY-FloorGeometryTolerance&&b[1]<=maxY+FloorGeometryTolerance;}).ToList();
            var outline=NativeWallContourRegion(walls,"outer",z);
            var relevant=outline.Components().Select(c=>new PlanarRegion(c)).Where(c=>!RegionIntersection(c,footprints).IsEmpty).ToList();
            if(relevant.Count==0)throw new InvalidOperationException("Помещения не попали в замкнутый внешний контур надстройки.");
            foreach(var component in relevant)result=RegionUnion(result,component);
            if(RegionDifference(footprints,result).Area*.09290304>.001)throw new InvalidOperationException("Внешний контур не охватывает все помещения надстройки.");
            return result;
        }
        private static PlanarRegion FillProjectionHoles(PlanarRegion region)
        {return region.Paths.Select(p=>PlanarRegion.FromRings(new[]{p.Select(v=>new[]{v.X/PlanarRegion.Scale,v.Y/PlanarRegion.Scale})})).Aggregate(PlanarRegion.Empty,RegionUnion);}
        public static string LevelGeometryScopeKey(IEnumerable<string> keys)
        {
            // An edited aggregate belongs to its original set of objects, never to a new classification.
            var ordered=keys.Distinct().OrderBy(k=>k,StringComparer.Ordinal).Select(k=>k.Length+":"+k);
            using(var hash=System.Security.Cryptography.SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(System.Text.Encoding.UTF8.GetBytes(string.Join("",ordered)))).Replace("-","");
        }
        private void PrepareRoofRatio(LevelSetting level,List<Record> scope,List<Tuple<Source,HostObject>> hosts,LevelGeometryReview review)
        {
            if(scope.Any(r=>r.Role=="unknown"))throw new InvalidOperationException("Распределите назначения надстройки: состав технических помещений не определён.");
            var footprint=scope.Select(RoomNetRegion).Aggregate(PlanarRegion.Empty,RegionUnion);
            var technical=scope.Where(r=>OneOf(r.Role,"technical-room","technical-space","technical-void","roof-vent")).ToList();
            var numerator=technical.Select(RoomNetRegion).Aggregate(PlanarRegion.Empty,RegionUnion);
            if(!numerator.IsEmpty)numerator=ReviewRegion("normative/technical/"+level.Key+"/"+LevelGeometryScopeKey(technical.Select(r=>r.Key)),scope[0],"Надстройка - технические помещения",level.Elevation,()=>numerator);
            review.TechnicalArea=numerator.Area*.09290304;
            var failures=new List<string>();
            foreach(var pair in hosts.Where(p=>p.Item2 is RoofBase||p.Item2 is Floor))
            {
                var source=pair.Item1;var element=pair.Item2;var box=ElementBounds(element,source.Transform);
                var anchor=element.Document.GetElement(element.LevelId) as Level;
                double elevation=anchor==null?double.NaN:source.Transform.OfPoint(new XYZ(0,0,anchor.ProjectElevation)).Z;
                if(box==null||box[2]>level.Elevation+1e-6||box[5]<level.Elevation-1e-6&&Math.Abs(elevation-level.Elevation)>1e-6)continue;
                if(RegionIntersection(BoundsRegion(box),footprint).IsEmpty)continue;
                string identity=source.Key+"/"+element.UniqueId;string name=source.Name+" / "+element.Name+" / ID "+IDHelper.ElIdValue(element.Id);
                try
                {
                    var region=Solids(element).Select(s=>ProjectionRegion(SolidUtils.CreateTransformed(s,source.Transform))).Aggregate(PlanarRegion.Empty,RegionUnion);
                    if(RegionIntersection(FillProjectionHoles(region),footprint).IsEmpty)continue;
                    review.Roofs.Add(new LevelRoofOption{Key=identity,Name=name,Area=region.Area*.09290304,Region=region,
                        Record=new Record{Source=source,Element=element,Level=level,Role="roof"}});
                }
                catch(OperationCanceledException){throw;}catch(Exception ex){failures.Add(name+": "+ex.Message);}
            }
            var selected=level.RoofSources??new List<string>();
            var roofs=selected.Count>0?review.Roofs.Where(r=>selected.Contains(r.Key)).ToList():review.Roofs;
            if(selected.Count>0&&roofs.Count!=selected.Distinct().Count())throw new InvalidOperationException("Выбранная кровля удалена, перемещена или больше не соответствует этажу. Выберите источники заново.");
            if(selected.Count==0&&(roofs.Count!=1||failures.Count>0))throw new InvalidOperationException("Нужно выбрать кровлю: найдено вариантов "+roofs.Count+". "+string.Join(" ",failures));
            if(roofs.Count==0)throw new InvalidOperationException("Кровля не выбрана.");
            var denominator=PlanarRegion.Empty;
            foreach(var roof in roofs)
            {
                var region=ReviewRegion("normative/roof/"+level.Key+"/"+roof.Key,roof.Record,"Кровля - площадь в плане",level.Elevation,()=>roof.Region);
                denominator=RegionUnion(denominator,region);review.Details.Add("Кровля: "+roof.Name);
            }
            review.RoofArea=denominator.Area*.09290304;
            if(RegionDifference(footprint,FillProjectionHoles(denominator)).Area*.09290304>.001)
                throw new InvalidOperationException("Выбранная кровля не охватывает надстройку в плане. Проверьте источники и 2D-границы.");
            review.Ratio=ValidRoofRatio(review.TechnicalArea.Value,review.RoofArea.Value);
        }
    }
}
}
