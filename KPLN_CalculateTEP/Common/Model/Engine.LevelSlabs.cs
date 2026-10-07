using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public class TopSlabOption
        {
            public string Key {get;set;}
            public string Elements {get;set;}
            public double InternalElevation {get;set;}
            public double AbsoluteMeters {get;set;}
            public double MissingArea {get;set;}=double.NaN;
            public bool CoversRooms {get{return MissingArea>=0&&MissingArea<=.001;}}
            public string Display {get{return AbsoluteMeters.ToString("0.000",CultureInfo.CurrentCulture)+" м | "+Elements+
                (CoversRooms?" | перекрывает все помещения":" | не перекрыто "+MissingArea.ToString("0.###")+" м²");}}
            internal PlanarRegion Region=PlanarRegion.Empty;
            internal List<string> ElementKeys=new List<string>();
        }
        public class TopSlabReview
        {
            public List<TopSlabOption> Options {get;set;}=new List<TopSlabOption>();
            public List<string> Problems {get;set;}=new List<string>();
            public TopSlabOption Selected {get;set;}
            public string Error {get;set;}
        }
        public partial class Engine
        {
            private readonly Dictionary<string,TopSlabReview> topSlabReviews=new Dictionary<string,TopSlabReview>();
            public TopSlabReview TopSlabReviewFor(LevelSetting level)
            {TopSlabReview result;return topSlabReviews.TryGetValue(level.Key,out result)?result:new TopSlabReview{Error="Перекрытия ещё не проверены."};}

            // All calculation geometry is in host internal feet. Absolute elevations are presentation/input only.
            public static double AbsoluteToInternal(double meters,double sharedOriginFeet) {return meters/.3048-sharedOriginFeet;}
            private double AbsoluteDatumToInternal(double meters)
            {return AbsoluteToInternal(meters,doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero).Elevation);}

            public static TopSlabOption SelectTopSlab(TopSlabReview review,string selectedKey)
            {
                if(!string.IsNullOrWhiteSpace(selectedKey))
                    return review.Options.FirstOrDefault(o=>o.Key==selectedKey&&o.CoversRooms);
                return review.Problems.Count==0&&review.Options.Count==1&&review.Options[0].CoversRooms?review.Options[0]:null;
            }
            private double ResolvedTopSlab(LevelSetting level)
            {
                if(level.TopSlabMode=="manual")return AbsoluteDatumToInternal(RequiredNumber(level.TopSlab,"абсолютная отметка верха перекрытия цоколя «"+level.Name+"», м"));
                var review=TopSlabReviewFor(level);
                var selected=SelectTopSlab(review,level.TopSlabChoice);
                if(selected==null)throw new InvalidOperationException("Верх перекрытия цоколя «"+level.Name+"»: "+review.Error+" Откройте выбор перекрытия в таблице этажей.");
                return selected.InternalElevation;
            }
            private string TopSlabRuleKey(LevelSetting level)
            {
                if(level.Kind!="basement"||level.Above!="auto")return "";
                if(level.TopSlabMode=="manual")
                {double value;return TryNumber(level.TopSlab,out value)?AbsoluteDatumToInternal(value).ToString("R",CultureInfo.InvariantCulture):"manual-invalid/"+level.TopSlab;}
                var selected=SelectTopSlab(TopSlabReviewFor(level),level.TopSlabChoice);
                return selected==null?"unresolved/"+level.Key:selected.InternalElevation.ToString("R",CultureInfo.InvariantCulture);
            }
            public void RefreshTopSlabs(Action<string> progress=null)
            {
                topSlabReviews.Clear();
                phaseCache.Clear();
                var levels=Config.Levels.Where(l=>l.Include&&l.Kind=="basement"&&l.Above=="auto").ToList();
                if(levels.Count==0)return;
                var sources=Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode!="exclude").ToList();
                // A scan owns its cache: edited geometry, moved links and a changed phase never reuse old elevations.
                var geometry=new Dictionary<string,Tuple<double,PlanarRegion,string>>();
                foreach(var level in levels)
                {
                    progress?.Invoke("Верх перекрытия цоколя: "+level.Source+" / "+level.Name);
                    var review=new TopSlabReview();topSlabReviews[level.Key]=review;
                    if(level.TopSlabMode=="manual")
                    {level.TopSlabSummary="Вручную: "+level.TopSlab+" м";continue;}
                    try {FindTopSlabs(level,review,sources,geometry,progress);}
                    catch(OperationCanceledException){throw;}
                    catch(Exception ex){review.Options.Clear();review.Problems.Add(ex.Message);}
                    review.Selected=SelectTopSlab(review,level.TopSlabChoice);
                    if(review.Selected!=null)
                    {
                        level.TopSlabSummary=(string.IsNullOrEmpty(level.TopSlabChoice)?"Авто: ":"Выбрано: ")+review.Selected.AbsoluteMeters.ToString("0.000")+" м";
                        // Old numeric values must not silently become a fallback after geometry disappears.
                        level.TopSlab=null;
                    }
                    else
                    {
                        review.Error=!string.IsNullOrEmpty(level.TopSlabChoice)?"Выбранный ранее вариант отсутствует или больше не перекрывает все помещения.":
                            review.Options.Count==0?"Подходящее перекрытие не найдено.":
                            review.Options.Any(o=>o.CoversRooms)?"Нужно подтвердить вариант перекрытия.":"Нет единой отметки перекрытия над всеми помещениями этажа.";
                        if(review.Problems.Count>0)review.Error+=" "+string.Join(" ",review.Problems.Distinct());
                        level.TopSlabSummary="Задать";
                    }
                }
            }
            private void FindTopSlabs(LevelSetting level,TopSlabReview review,List<Source> sources,
                Dictionary<string,Tuple<double,PlanarRegion,string>> geometry,Action<string> progress)
            {
                var source=sources.SingleOrDefault(s=>s.Key==level.Key.Substring(0,level.Key.LastIndexOf('/')));
                if(source==null)throw new InvalidOperationException("Источник уровня недоступен.");
                var rooms=source.Elements.OfType<Room>().Where(r=>IsPlacedRoom(r)&&r.Area>1e-9&&PhaseAccepted(source,r)).ToList();
                var onLevel=rooms.Where(r=>source.Key+"/"+r.Level.UniqueId==level.Key).ToList();
                if(onLevel.Count==0)throw new InvalidOperationException("На уровне нет размещённых помещений с положительной площадью.");
                var footprint=PlanarRegion.Empty;
                foreach(var room in onLevel)
                {
                    using(var options=new SpatialElementBoundaryOptions{SpatialElementBoundaryLocation=SpatialElementBoundaryLocation.Finish})
                    {
                        var loops=room.GetBoundarySegments(options);
                        if(loops==null||loops.Count==0)throw new InvalidOperationException("Не прочитан контур помещения ID "+IDHelper.ElIdValue(room.Id)+".");
                        var rings=new List<List<double[]>>();
                        foreach(var loop in loops)
                        {
                            var ring=new List<double[]>();
                            foreach(var segment in loop)
                            {
                                var points=CurvePoints(segment.GetCurve().CreateTransformed(source.Transform),ArcChordTolerance);
                                ring.AddRange(points.Take(points.Count-1).Select(p=>new[]{p[0],p[1]}));
                            }
                            rings.Add(ring);
                        }
                        footprint=PlanarRegion.Combine(footprint,PlanarRegion.FromRings(rings),CT.ctUnion);
                    }
                }
                // Bound the search to the next occupied level in this source, not an arbitrary metre offset.
                var higher=rooms.Select(r=>source.Transform.OfPoint(new XYZ(0,0,r.Level.ProjectElevation)).Z)
                    .Where(z=>z>level.Elevation+1e-6).ToList();
                if(higher.Count==0)
                    higher=onLevel.Where(r=>r.UpperLimit!=null).Select(r=>source.Transform.OfPoint(new XYZ(0,0,r.UpperLimit.ProjectElevation+r.LimitOffset)).Z)
                        .Where(z=>z>level.Elevation+1e-6).ToList();
                if(higher.Count==0)throw new InvalidOperationException("Не определена верхняя граница поиска: нет следующего занятого уровня или верхней границы помещений.");
                double ceiling=higher.Min();
                var bounds=onLevel.Select(r=>ElementBounds(r,source.Transform)).Where(b=>b!=null).ToList();
                if(bounds.Count!=onLevel.Count)throw new InvalidOperationException("Не прочитаны габариты помещений для поиска перекрытий.");
                var extent=new[]{bounds.Min(b=>b[0]),bounds.Min(b=>b[1]),bounds.Max(b=>b[3]),bounds.Max(b=>b[4])};
                int count=0;
                foreach(var floorSource in sources)
                    foreach(var floor in floorSource.Elements.OfType<Floor>())
                    {
                        if(++count%50==0)progress?.Invoke("Поиск перекрытий над «"+level.Name+"»: проверено "+count);
                        if(!PhaseAccepted(floorSource,floor))continue;
                        var box=ElementBounds(floor,floorSource.Transform);
                        if(box==null||box[3]<extent[0]||box[0]>extent[2]||box[4]<extent[1]||box[1]>extent[3]||box[2]<=level.Elevation+1e-6)continue;
                        var floorLevel=floor.Document.GetElement(floor.LevelId) as Level;
                        double anchor=floorLevel==null?double.NaN:floorSource.Transform.OfPoint(new XYZ(0,0,floorLevel.ProjectElevation)).Z;
                        if(box[2]>ceiling+1e-6&&Math.Abs(anchor-ceiling)>1e-6)continue;
                        string key=floorSource.Key+"/"+floor.UniqueId;
                        string label=floorSource.Name+" / "+floor.Name+" / ID "+IDHelper.ElIdValue(floor.Id);
                        Tuple<double,PlanarRegion,string> item;
                        if(!geometry.TryGetValue(key,out item))
                        {
                            try
                            {
                                var faces=HostObjectUtils.GetTopFaces(floor).Select(r=>floor.GetGeometryObjectFromReference(r) as Face).ToList();
                                if(faces.Count==0)throw new InvalidOperationException("Нет верхних граней.");
                                double? top=null;var region=PlanarRegion.Empty;
                                foreach(var face in faces)
                                {
                                    var plane=face as PlanarFace;
                                    var normal=plane==null?null:floorSource.Transform.OfVector(plane.FaceNormal).Normalize();
                                    if(normal==null||Math.Abs(normal.X)>1e-10||Math.Abs(normal.Y)>1e-10)
                                        throw new InvalidOperationException("Верх наклонный или криволинейный; единой отметки нет.");
                                    double z=floorSource.Transform.OfPoint(plane.Origin).Z;
                                    if(top.HasValue&&Math.Abs(top.Value-z)>1e-6)throw new InvalidOperationException("Верх перекрытия имеет несколько отметок.");
                                    top=z;bool curved=false;var rings=new List<List<double[]>>();
                                    foreach(var loop in face.GetEdgesAsCurveLoops())
                                    {
                                        var transformed=new CurveLoop();foreach(var curve in loop)transformed.Append(curve.CreateTransformed(floorSource.Transform));
                                        rings.Add(PolygonRing(transformed,ref curved));
                                    }
                                    region=PlanarRegion.Combine(region,PlanarRegion.FromRings(rings),CT.ctUnion);
                                }
                                item=Tuple.Create(top.Value,region,(string)null);
                            }
                            catch(OperationCanceledException){throw;}
                            catch(Exception ex){item=Tuple.Create(double.NaN,PlanarRegion.Empty,ex.Message);}
                            geometry[key]=item;
                        }
                        if(item.Item3!=null){review.Problems.Add(label+": "+item.Item3);continue;}
                        if(PlanarRegion.Combine(footprint,item.Item2,CT.ctIntersection).IsEmpty)continue;
                        if(double.IsNaN(anchor)||anchor<=level.Elevation+1e-6)
                            review.Problems.Add(label+": перекрытие привязано к цоколю, нижнему уровню или не имеет уровня. Подтвердите, что это перекрытие НАД цоколем, а не его пол со смещением.");
                        var candidate=review.Options.FirstOrDefault(o=>Math.Abs(o.InternalElevation-item.Item1)<=1e-6);
                        if(candidate==null)
                        {candidate=new TopSlabOption{InternalElevation=item.Item1,AbsoluteMeters=doc.ActiveProjectLocation.GetProjectPosition(new XYZ(0,0,item.Item1)).Elevation*.3048};review.Options.Add(candidate);}
                        candidate.ElementKeys.Add(key);candidate.Elements=string.IsNullOrEmpty(candidate.Elements)?label:candidate.Elements+"; "+label;
                        candidate.Region=PlanarRegion.Combine(candidate.Region,item.Item2,CT.ctUnion);
                    }
                FinalizeTopSlabOptions(review,footprint);
            }
            private static void FinalizeTopSlabOptions(TopSlabReview review,PlanarRegion footprint)
            {
                foreach(var option in review.Options)
                {
                    option.Key=string.Join("|",option.ElementKeys.OrderBy(k=>k,StringComparer.Ordinal));
                    option.MissingArea=PlanarRegion.Combine(footprint,option.Region,CT.ctDifference).Area*.09290304;
                }
                review.Options=review.Options.OrderBy(o=>o.InternalElevation).ToList();
            }
        }
    }
}
