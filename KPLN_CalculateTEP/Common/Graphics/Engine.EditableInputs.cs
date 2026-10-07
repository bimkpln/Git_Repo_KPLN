using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class ReviewBinding
        {
            public string Key { get; set; }
            public string RegionId { get; set; }
            public double Elevation { get; set; }
            public string Placement { get; set; }
            public List<List<double[]>> InitialBoundary { get; set; }
            public bool RequiresReview {get;set;}
            public string DraftReason {get;set;}
            public string DraftDescription {get;set;}
        }
        public class ReviewInput
        {
            public string Key { get; set; }
            public string Kind { get; set; }
            public string DisplayKind {get{return Kind=="Контур этажа - exterior"?"Контур этажа - внешняя граница":Kind=="Контур этажа - interior"?"Контур этажа - внутренняя граница":Kind=="Контур этажа - core"?"Контур этажа - граница несущего слоя":Kind;}}
            public string Level { get; set; }
            public string Source { get; set; }
            public string SourceKey { get; set; }
            public string Apartment { get; set; }
            public string ObjectName {get;set;}
            public string Building {get;set;}
            public string Section {get;set;}
            public string RegionElementId {get;set;}
            public double? SavedArea {get;set;}
            public bool IsDraft {get;set;}
            public string DraftReason {get;set;}
            public string DraftDescription {get;set;}
            public string Element { get; set; }
            public double? ContourArea {get {return Region!=null&&!Region.IsEmpty?(double?)(Region.Area*.09290304):SavedArea>0?SavedArea:null;}}
            public double Area { get { return ContourArea??0; } }
            public string AreaDescription {get {return ContourArea.HasValue?"Площадь контура: "+ContourArea.Value.ToString("N3")+" м²":"Площадь контура: не определена";}}
            public bool CanOpen {get {return !string.IsNullOrEmpty(RegionElementId);}}
            public string ReviewState {get {return IsDraft?"Предварительный контур - не включён в расчёт":!CanOpen?"Область не построена":!ContourArea.HasValue?"Контур не рассчитан":"Расчётный контур";}}
            public double MeasurementElevationMeters { get { return Elevation * .3048; } set {Elevation=value/.3048;} }
            public double PlanElevationMeters {get {return PlanElevation*.3048;}set {PlanElevation=value/.3048;}}
            public string Status { get; set; }
            public bool ObjectOverlay { get; set; }
            public string Role { get; set; }
            internal PlanarRegion Region;
            internal double Elevation;
            internal double PlanElevation;
        }
        public partial class Engine
        {
            private readonly Dictionary<string, ReviewInput> reviewInputs = new Dictionary<string, ReviewInput>();
            private List<ReviewBinding> reviewBindings = new List<ReviewBinding>();
            private string reviewConfiguration;
            public List<ReviewInput> ReviewInputs { get { return reviewInputs.Values.OrderBy(r => r.Elevation).ThenBy(r => r.Kind).ThenBy(r => r.Element).ToList(); } }
            public string PendingReviewKey { get; set; }
            private string ReviewPlacement()
            {
                return string.Join(";", Sources.Where(s => s.Loaded && s.Mode == "include").OrderBy(s => s.Key).Select(s =>
                    s.Key + ":" + string.Join(",", new[] {s.Transform.Origin, s.Transform.BasisX, s.Transform.BasisY, s.Transform.BasisZ}
                        .SelectMany(p => new[] {p.X,p.Y,p.Z}).Select(x => x.ToString("R", CultureInfo.InvariantCulture)))));
            }
            private void BeginReviewInputs()
            {
                reviewConfiguration=Serialize(Config);reviewInputs.Clear();editedObjectOverlays.Clear();
                automaticDimensions.Clear();automaticDimensionErrors.Clear();clearanceGeometry.Clear();clearanceErrors.Clear();clearanceFaces.Clear();automaticGeometryRecords=new List<Record>();contourRoomRecords.Clear();
                roomFloorRegions.Clear();roomFloorFailures.Clear();roomFloorWarnings.Clear();roomNetRegions.Clear();roomCentreRegions.Clear();measuredRoomRegions.Clear();shapes.Clear();
                reviewBindings = Read<List<ReviewBinding>>("editable-inputs") ?? new List<ReviewBinding>();
            }
            private PlanarRegion ReviewRegion(string key, Record record, string kind, double elevation, Func<PlanarRegion> automatic)
            {
                key = Config.Method + "/" + (Config.Phase ?? "") + "/" + key;
                ReviewInput input;
                if (!reviewInputs.TryGetValue(key, out input))
                    reviewInputs[key] = input = new ReviewInput { Key=key, Kind=kind, Level=record.Level?.Name, Source=record.Source.Name,SourceKey=record.Source.Key,Apartment=record.Apartment,ObjectName=record.Element.Name,Building=record.Building,Section=record.Section,
                        Element=IDHelper.ElIdValue(record.Element.Id).ToString(), Elevation=elevation, PlanElevation=kind.StartsWith("Застройка")?elevation:record.Z,Role=record.Role };
                try
                {
                    var binding = reviewBindings.SingleOrDefault(b => b.Key == key);
                    if ((Config.UseReviewRegions||automaticAreaInputs) && binding != null)
                    {
                        if (binding.Placement != ReviewPlacement()) throw new InvalidOperationException("Изменился состав или размещение источников. Перепривяжите проверенную область; прежняя привязка не применяется.");
                        var region = doc.GetElement(binding.RegionId) as FilledRegion;
                        if (region == null) throw new InvalidOperationException("Расчётная область удалена. Привяжите новую область к этой строке.");
                        var view = doc.GetElement(region.OwnerViewId) as ViewPlan;
                        if (view?.GenLevel == null || Math.Abs(view.GenLevel.ProjectElevation-input.PlanElevation)>FloorGeometryTolerance || Math.Abs(binding.Elevation-elevation)>FloorGeometryTolerance)
                            throw new InvalidOperationException("План области или отметка расчётного контура изменились. Требуется план уровня «"+input.Level+"» с отметкой "+input.PlanElevationMeters.ToString("0.###")+" м; отметка контура "+input.MeasurementElevationMeters.ToString("0.###")+" м.");
                        input.Region = RegionFromLoops(region.GetBoundaries(),true);
                        if(input.Region.IsEmpty)throw new InvalidOperationException("У расчётной области нет замкнутой границы.");
                        input.RegionElementId=IDHelper.ElIdValue(region.Id).ToString();
                        input.IsDraft=binding.RequiresReview;input.DraftReason=binding.DraftReason;input.DraftDescription=binding.DraftDescription;
                        bool edited=ValidateReviewBinding(binding,input.Region,automatic);
                        input.IsDraft=false;input.DraftReason=null;input.DraftDescription=null;
                        input.Status = "Из редактируемой области ID " + IDHelper.ElIdValue(region.Id);
                        if(edited)record.Manual=true;
                        Notice("EDITABLE_INPUT", edited?"Предупреждение":"Информация", kind+": площадь прочитана из расчётной области ID "+IDHelper.ElIdValue(region.Id)+(edited?". Учтена пользовательская граница.":"."),record);
                    }
                    else
                    {
                        if(automaticAreaInputs&&areaBuildFailures.TryGetValue(key,out var buildFailure))throw new InvalidOperationException(buildFailure);
                        input.Region=automatic();
                        if(input.Region==null||input.Region.IsEmpty)throw new InvalidOperationException("Не получен замкнутый контур «"+kind+"» на отметке "+input.MeasurementElevationMeters.ToString("0.###")+" м. Проверьте геометрию и границы исходного объекта.");
                        input.IsDraft=false;input.DraftReason=null;input.DraftDescription=null;
                        input.Status = binding == null ? "Автоматический контур" : "Исправления отключены";
                    }
                    if(automaticAreaInputs&&binding==null)
                    {
                        if(areaBuildFailures.TryGetValue(key,out var failure))throw new InvalidOperationException(failure);
                        CreateEditableInputs(reportProgress,new[]{input});
                        binding=reviewBindings.SingleOrDefault(b=>b.Key==key);
                        if(binding==null){areaBuildFailures[key]=input.Status;throw new InvalidOperationException(input.Status);}
                        var created=doc.GetElement(binding.RegionId) as FilledRegion;
                        input.Region=RegionFromLoops(created.GetBoundaries(),true);
                        if(input.Region.IsEmpty)throw new InvalidOperationException("Созданная расчётная область пуста.");
                    }
                    return input.Region;
                }
                catch (OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { throw; }
                catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                catch (Exception ex) { PreserveFailedReviewInput(input,ex);throw; }
            }
            internal static bool ValidateReviewBinding(ReviewBinding binding,PlanarRegion actual,Func<PlanarRegion> automatic)
            {
                bool changed=binding.InitialBoundary==null||OverlayBoundaryChanged(PlanarRegion.FromRings(binding.InitialBoundary),actual);
                // An untouched preliminary outline must pass its original geometry/coverage checks.
                // Editing or explicitly rebinding it is the existing user correction workflow.
                if(binding.RequiresReview&&(binding.InitialBoundary==null||!changed))ValidateCreatedArea(automatic(),actual);
                return changed;
            }
            internal static void PreserveFailedReviewInput(ReviewInput input,Exception error)
            {
                var provisional=error as UnconfirmedEnvelopeException;
                if(provisional?.Candidate!=null&&!provisional.Candidate.IsEmpty)
                {if(input.Region==null||!input.CanOpen)input.Region=provisional.Candidate;input.DraftDescription=provisional.CandidateDescription;}
                input.SavedArea=null;
                input.IsDraft=input.Region!=null&&!input.Region.IsEmpty;
                input.DraftReason=error.Message;
                input.Status=input.IsDraft?"Не включён в расчёт. "+error.Message:error.Message;
            }
            private void SaveReviewBindings()
            {
                var data = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).FirstOrDefault(e => Owned(e) && Kind(e)=="editable-inputs") as DataStorage;
                if(data==null){data=DataStorage.Create(doc);data.Name=Owner+"/editable-inputs";}
                Tag(data,"editable-inputs",PackStorage(Serialize(reviewBindings)));
            }
            private List<ReviewInput> SnapshotReviewInputs()
            {
                foreach(var input in ReviewInputs)
                {
                    input.SavedArea=input.ContourArea;
                    var binding=reviewBindings.FirstOrDefault(b=>b.Key==input.Key);
                    var region=binding==null?null:doc.GetElement(binding.RegionId) as FilledRegion;
                    input.RegionElementId=region==null?null:IDHelper.ElIdValue(region.Id).ToString();
                }
                return ReviewInputs;
            }
            private void RestoreReviewInputs()
            {
                reviewInputs.Clear();
                reviewBindings=Read<List<ReviewBinding>>("editable-inputs")??new List<ReviewBinding>();
                foreach(var input in Last?.ReviewAreas??new List<ReviewInput>())
                {
                    if(input==null||string.IsNullOrEmpty(input.Key))continue;
                    var binding=reviewBindings.FirstOrDefault(b=>b.Key==input.Key);
                    var region=binding==null?null:doc.GetElement(binding.RegionId) as FilledRegion;
                    string id=region==null?null:IDHelper.ElIdValue(region.Id).ToString();
                    if(input.RegionElementId!=id)input.Status=id==null?"Расчётная область отсутствует. Привяжите другую область и пересчитайте.":"Привязана другая область ID "+id+". Пересчитайте ТЭП для обновления площади.";
                    input.RegionElementId=id;reviewInputs[input.Key]=input;
                }
            }
            private IList<CurveLoop> ReviewLoops(PlanarRegion region, double elevation)
            {
                return region.Paths.Select(path => CurveLoop.Create(path.Select((p,i) => (Curve)Line.CreateBound(
                    new XYZ(p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale,elevation),
                    new XYZ(path[(i+1)%path.Count].X/PlanarRegion.Scale,path[(i+1)%path.Count].Y/PlanarRegion.Scale,elevation))).ToList())).ToList();
            }
            internal static void ValidateCreatedArea(PlanarRegion expected,PlanarRegion actual)
            {
                double before=(expected?.Area??0)*.09290304,after=(actual?.Area??0)*.09290304;
                double difference=Math.Abs(before-after);
                double displaced=actual==null||expected==null?double.PositiveInfinity:
                    (RegionDifference(expected,actual).Area+RegionDifference(actual,expected).Area)*.09290304;
                if(actual==null||actual.IsEmpty||difference>.001||displaced>.001)
                    throw new InvalidOperationException("Не подтверждено совпадение области с исходным контуром: исходная площадь "+before.ToString("0.######",CultureInfo.InvariantCulture)+
                        " м²; после создания "+after.ToString("0.######",CultureInfo.InvariantCulture)+" м²; разница "+difference.ToString("0.######",CultureInfo.InvariantCulture)+
                        " м²; различие контуров "+displaced.ToString("0.######",CultureInfo.InvariantCulture)+" м²; границ "+(actual?.Paths.Count??0)+". Допуск 0,001 м² не увеличен.");
            }
            public int CreateEditableInputs(Action<string> progress=null,IEnumerable<ReviewInput> requested=null)
            {
                var candidates=(requested??ReviewInputs).Where(r=>r.Region!=null&&!r.Region.IsEmpty&&!reviewBindings.Any(b=>b.Key==r.Key)).ToList();
                if(candidates.Count==0)return 0;
                var report=current??Last??new Run();var previous=reviewBindings.ToList();int created=0;
                int issueStart=report.Issues.Count;
                try
                {
                    Transaction("ТЭП: расчётные 2D-области",report,()=>
                    {
                        var type=new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault();
                        var planType=new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(v=>v.ViewFamily==ViewFamily.FloorPlan);
                        if(type==null||planType==null)throw new InvalidOperationException("Нужны тип цветовой области и тип плана этажа в проекте.");
                        var fill=new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Cast<FillPatternElement>().FirstOrDefault(f=>f.GetFillPattern().IsSolidFill);
                        if(fill!=null)
                        {
                            var ownedType=new FilteredElementCollector(doc).OfClass(typeof(FilledRegionType)).Cast<FilledRegionType>().FirstOrDefault(t=>Owned(t)&&Kind(t)=="review-type");
                            if(ownedType==null){ownedType=(FilledRegionType)type.Duplicate("ТЭП_РасчётныеОбласти_"+Guid.NewGuid().ToString("N").Substring(0,8));Tag(ownedType,"review-type");}
                            ownedType.IsMasking=false;ownedType.ForegroundPatternId=fill.Id;type=ownedType;
                        }
                        foreach(var group in candidates.GroupBy(r=>new {r.PlanElevation,Layer=AreaViewLayer(r)}))
                        {
                            var level=new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault(l=>Math.Abs(l.ProjectElevation-group.Key.PlanElevation)<FloorGeometryTolerance);
                            if(level==null){level=Level.Create(doc,group.Key.PlanElevation);level.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY)?.Set(0);Tag(level,"review-level");}
                            string viewKey=Config.Method+"/"+(Config.Phase??"")+"/"+group.Key.PlanElevation.ToString("R",CultureInfo.InvariantCulture)+"/"+group.Key.Layer;
                            var view=new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().FirstOrDefault(v=>Owned(v)&&Kind(v)=="review-overlay-plan"&&v.GenLevel?.Id==level.Id&&v.GetEntity(StorageSchema()).Get<string>(StorageSchema().GetField("Payload"))==viewKey);
                            if(view==null){view=ViewPlan.Create(doc,planType.Id,level.Id);view.Name=UniqueName("ТЭП_"+group.Key.Layer+"_"+group.First().Level+"_"+Config.Method+"_"+(Config.Phase??""));Tag(view,"review-overlay-plan",viewKey);}
                            var pending=new List<Tuple<ReviewInput,FilledRegion>>();
                            foreach(var input in group)
                            {
                                progress?.Invoke("Построение 2D: "+input.Level+"; "+input.Kind+"; ID "+input.Element);
                                using(var sub=new SubTransaction(doc))
                                {
                                    sub.Start();
                                    try
                                    {
                                        var region=FilledRegion.Create(doc,type.Id,view.Id,ReviewLoops(input.Region,input.PlanElevation));
                                        Tag(region,"review-input",input.Key);
                                        var colour=OverlayColour(input.IsDraft?"unknown":input.Role);
                                        var style=new OverrideGraphicSettings().SetProjectionLineColor(colour).SetSurfaceTransparency(65);
                                        if(fill!=null)style.SetSurfaceForegroundPatternId(fill.Id).SetSurfaceForegroundPatternColor(colour);
                                        if(input.Kind.StartsWith("Контур этажа")&&!input.IsDraft)style.SetSurfaceForegroundPatternVisible(false).SetSurfaceBackgroundPatternVisible(false);
                                        view.SetElementOverrides(region.Id,style);
                                        region.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)?.Set((input.IsDraft?"Предварительный контур. "+input.DraftDescription+" | ":"")+input.Kind+" | "+input.Source+" | ID "+input.Element);
                                        if(sub.Commit()!=TransactionStatus.Committed)throw new InvalidOperationException("Создание области отменено Revit.");
                                        pending.Add(Tuple.Create(input,region));
                                    }
                                    catch(OperationCanceledException){throw;}
                                    catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                    catch(Exception ex){if(sub.GetStatus()==TransactionStatus.Started)sub.RollBack();input.Status="Не построено: "+ex.Message;report.Issue("EDITABLE_CREATE","Предупреждение",input.Status,source:input.Source,element:input.Element);}
                                }
                            }
                            // Manual regeneration mode: new sketch boundaries are not readable before this.
                            // One update per floor/layer avoids regenerating the entire model for each room.
                            progress?.Invoke("Обновление геометрии 2D: "+group.First().Level+"; "+pending.Count+" областей");
                            doc.Regenerate();
                            var rejected=new List<ElementId>();
                            foreach(var pair in pending)
                            {
                                var input=pair.Item1;var region=pair.Item2;
                                progress?.Invoke("Проверка созданной области: ID "+input.Element);
                                try
                                {
                                    var roundtrip=RegionFromLoops(region.GetBoundaries(),true);ValidateCreatedArea(input.Region,roundtrip);
                                    reviewBindings.Add(new ReviewBinding{Key=input.Key,RegionId=region.UniqueId,Elevation=input.Elevation,Placement=ReviewPlacement(),InitialBoundary=roundtrip.Paths.Select(p=>p.Select(v=>new[]{v.X/PlanarRegion.Scale,v.Y/PlanarRegion.Scale}).ToList()).ToList(),RequiresReview=input.IsDraft,DraftReason=input.DraftReason,DraftDescription=input.DraftDescription});
                                    input.Region=roundtrip;input.RegionElementId=IDHelper.ElIdValue(region.Id).ToString();
                                    input.Status=input.IsDraft?"Предварительная область ID "+input.RegionElementId+" создана для исправления. Не включена в расчёт: "+input.DraftReason:"Расчётная область ID "+input.RegionElementId;created++;
                                }
                                catch(OperationCanceledException){throw;}
                                catch(Autodesk.Revit.Exceptions.RegenerationFailedException){throw;}
                                catch(Exception ex){rejected.Add(region.Id);input.Status="Не построено: "+ex.Message;report.Issue("EDITABLE_CREATE","Предупреждение",input.Status,source:input.Source,element:input.Element);}
                            }
                            if(rejected.Count>0)doc.Delete(rejected);
                        }
                        SaveReviewBindings();
                    });
                }
                catch(Autodesk.Revit.Exceptions.RegenerationFailedException ex)
                {reviewBindings=previous;throw new OperationCanceledException("Revit не смог обновить геометрию. Транзакция построения отменена.",ex);}
                catch{reviewBindings=previous;throw;}
                OverlayIssues=OverlayIssues.Concat(report.Issues.Skip(issueStart).Where(i=>i.Code=="EDITABLE_CREATE")).ToList();
                return created;
            }
            private static string AreaViewLayer(ReviewInput input)
            {
                if(input.Kind.StartsWith("Надстройка"))return "Надстройки";
                if(input.Kind.StartsWith("Кровля"))return "Кровля";
                if(input.Kind.StartsWith("Контур этажа")||input.Kind.Contains("по осям"))return "СПП";
                if(input.Kind.StartsWith("Застройка"))return "Застройка";
                return "Помещения_z="+(input.Elevation*.3048).ToString("0.###",CultureInfo.InvariantCulture);
            }
            public void RequestReviewNavigation(ReviewInput input)
            {
                var binding=reviewBindings.SingleOrDefault(b=>b.Key==input.Key);
                var region=binding==null?null:doc.GetElement(binding.RegionId) as FilledRegion;
                if(region==null)throw new InvalidOperationException("Область ещё не создана. Создайте доступные области или привяжите нарисованную вручную.");
                RequestedDetail=new Detail{SourceKey="host",Element=IDHelper.ElIdValue(region.Id).ToString(),ViewId=doc.GetElement(region.OwnerViewId).UniqueId};
            }
            private void BindReviewSelection()
            {
                ReviewInput input;
                if(!reviewInputs.TryGetValue(PendingReviewKey??"",out input))throw new InvalidOperationException("Выберите расчётный контур в списке.");
                var picked=ui.Selection.PickObject(ObjectType.Element,new ReviewSelectionFilter(),"Выберите проверенную цветовую область на плане уровня «"+input.Level+"»");
                var region=(FilledRegion)doc.GetElement(picked);
                var level=(doc.GetElement(region.OwnerViewId) as ViewPlan)?.GenLevel;
                if(level==null||Math.Abs(level.ProjectElevation-input.PlanElevation)>FloorGeometryTolerance)throw new InvalidOperationException("Цветовая область должна быть на плане уровня «"+input.Level+"» (отметка плана "+input.PlanElevationMeters.ToString("0.###")+" м).");
                var shape=RegionFromLoops(region.GetBoundaries(),true);
                if(shape.IsEmpty)throw new InvalidOperationException("Выбрана пустая область.");
                if(reviewBindings.Any(b=>b.RegionId==region.UniqueId&&b.Key!=input.Key))throw new InvalidOperationException("Эта область уже назначена другому контуру. Создайте отдельную область.");
                var previous=reviewBindings.ToList();
                try
                {
                    Transaction("ТЭП: привязать расчётную область",Last??new Run(),()=>
                    {
                        reviewBindings.RemoveAll(b=>b.Key==input.Key);
                        reviewBindings.Add(new ReviewBinding{Key=input.Key,RegionId=region.UniqueId,Elevation=input.Elevation,Placement=ReviewPlacement()});
                        SaveReviewBindings();
                    });
                    input.RegionElementId=IDHelper.ElIdValue(region.Id).ToString();
                    input.Status="Привязана область ID "+input.RegionElementId+". Пересчитайте ТЭП для обновления площади.";
                }
                catch{reviewBindings=previous;throw;}
            }
            private sealed class ReviewSelectionFilter:ISelectionFilter
            {
                public bool AllowElement(Element e){return e is FilledRegion;}
                public bool AllowReference(Reference r,XYZ p){return false;}
            }
        }
    }
}
