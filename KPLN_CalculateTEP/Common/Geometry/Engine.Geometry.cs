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
            private static IEnumerable<Solid> GeometrySolids(GeometryElement geometry)
            {
                if(geometry==null)yield break;
                foreach(GeometryObject o in geometry)
                {var solid=o as Solid;if(solid!=null&&solid.Volume>1e-9)yield return solid;
                    var instance=o as GeometryInstance;if(instance!=null)foreach(var s in GeometrySolids(instance.GetInstanceGeometry()))yield return s;}
            }
            private static List<Solid> Solids(Element e)
            {return GeometrySolids(e.get_Geometry(new Options{DetailLevel=ViewDetailLevel.Fine,IncludeNonVisibleObjects=false})).ToList();}
            private static List<double[]> PolygonRing(CurveLoop loop,ref bool curved,double chordTolerance=ArcChordTolerance)
            {
                var points=new List<double[]>();
                foreach(var curve in loop)
                {
                    curved|=!(curve is Line);
                    var part=CurvePoints(curve,chordTolerance);
                    for(int i=0;i+1<part.Count;i++)points.Add(new[]{part[i][0],part[i][1]});
                    if(points.Count>200000)throw new InvalidOperationException("Контур превысил лимит разбиения при допуске 0,1 мм. Точность не снижена.");
                }
                return points;
            }
            private static PlanarRegion FaceRegion(PlanarFace face,ref bool curved)
            {
                var rings=new List<List<double[]>>();
                foreach(var loop in face.GetEdgesAsCurveLoops())rings.Add(PolygonRing(loop,ref curved));
                return PlanarRegion.FromRings(rings);
            }
            private sealed class HorizontalCaps
            {internal PlanarRegion Up=PlanarRegion.Empty,Down=PlanarRegion.Empty;}
            private static bool TryLayeredBody(Solid solid,out LayeredBody body,out string reason)
            {
                body=null;reason=null;if(solid==null){body=new LayeredBody();return true;}
                Tuple<LayeredBody,string> cached;
                if(planarBodies!=null&&planarBodies.TryGetValue(solid,out cached)){body=cached.Item1;reason=cached.Item2;return body!=null;}
                try
                {
                    planarCheckpoint?.Invoke();
                    if(solid.Volume<=0||double.IsNaN(solid.Volume)||double.IsInfinity(solid.Volume))throw new NotSupportedException("Исходное тело не имеет положительного конечного объёма.");
                    var caps=new SortedDictionary<double,HorizontalCaps>();bool curved=false;
                    foreach(Face face in solid.Faces)
                    {
                        var plane=face as PlanarFace;
                        if(plane!=null)
                        {
                            if(Math.Abs(plane.FaceNormal.Z)<1e-10)continue;
                            if(Math.Abs(Math.Abs(plane.FaceNormal.Z)-1)>1e-10)throw new NotSupportedException("Есть наклонная грань; точное разложение на вертикальные участки неприменимо.");
                            double z=Math.Round(plane.Origin.Z,8);HorizontalCaps cap;
                            if(!caps.TryGetValue(z,out cap))caps[z]=cap=new HorizontalCaps();
                            var region=FaceRegion(plane,ref curved);
                            if(plane.FaceNormal.Z>0)cap.Up=PlanarRegion.Combine(cap.Up,region,CT.ctUnion);
                            else cap.Down=PlanarRegion.Combine(cap.Down,region,CT.ctUnion);
                        }
                        else
                        {
                            var cylinder=face as CylindricalFace;
                            if(cylinder==null||Math.Abs(Math.Abs(cylinder.Axis.Normalize().Z)-1)>1e-10)
                                throw new NotSupportedException("Есть непризматическая криволинейная грань; используется объёмная геометрия Revit.");
                        }
                    }
                    if(caps.Count<2)throw new NotSupportedException("Не найдены горизонтальные границы высотных участков.");
                    var candidate=new LayeredBody{Curved=curved};var active=PlanarRegion.Empty;var levels=caps.Keys.ToList();
                    for(int i=0;i<levels.Count;i++)
                    {
                        planarCheckpoint?.Invoke();var cap=caps[levels[i]];
                        active=PlanarRegion.Combine(active,cap.Down,CT.ctUnion);
                        active=PlanarRegion.Combine(active,cap.Up,CT.ctDifference);
                        if(i+1<levels.Count&&!active.IsEmpty)candidate.Layers.Add(new PlanarLayer{Bottom=levels[i],Top=levels[i+1],Region=active});
                    }
                    if(!active.IsEmpty||candidate.Layers.Count==0)throw new NotSupportedException("Горизонтальные грани не образуют замкнутый набор вертикальных участков.");
                    double tolerance=Math.Max(1e-7,solid.Volume*1e-8)+candidate.Layers.Sum(l=>l.Region.Perimeter*(curved?ArcChordTolerance:2/PlanarRegion.Scale)*(l.Top-l.Bottom)*4+l.Region.Area*2e-8);
                    if(Math.Abs(candidate.Volume-solid.Volume)>tolerance)throw new NotSupportedException("Разложение не сохранило объём в пределах допуска; исходное тело не заменено приближённой призмой.");
                    body=candidate;
                }
                catch(System.OperationCanceledException){throw;}
                catch(Exception ex){reason=ex.Message;}
                if(planarBodies!=null)planarBodies[solid]=Tuple.Create(body,reason);
                return body!=null;
            }
            private static LayeredBody RequireLayeredBody(Solid solid)
            {LayeredBody body;string reason;if(!TryLayeredBody(solid,out body,out reason))throw new InvalidOperationException("Плоское вычитание: "+reason);return body;}
            private static Solid BuildPlanarSolid(PlanarRegion region,double bottom,double top,bool curved=false)
            {
                if(region==null||region.IsEmpty||top<=bottom)return null;
                double expected=region.Area*(top-bottom);Solid shape=null;var failures=new List<string>();
                // Revit's curve-construction tolerance is larger than its geometric precision.
                // Build small edges in local enlarged coordinates, then apply an exact uniform
                // inverse transform. No vertices are removed or shifted relative to each other.
                foreach(double scale in new[]{1.0,1000.0})
                {
                    try
                    {
                        planarCheckpoint?.Invoke();var origin=region.Paths[0][0];
                        double x=origin.X/PlanarRegion.Scale,y=origin.Y/PlanarRegion.Scale;
                        var loops=new List<CurveLoop>();
                        foreach(var path in region.Paths)
                        {
                            var curves=new List<Curve>();
                            for(int i=0;i<path.Count;i++)
                            {
                                var p=path[i];var q=path[(i+1)%path.Count];
                                curves.Add(Line.CreateBound(new XYZ((p.X-origin.X)/PlanarRegion.Scale*scale,(p.Y-origin.Y)/PlanarRegion.Scale*scale,0),
                                    new XYZ((q.X-origin.X)/PlanarRegion.Scale*scale,(q.Y-origin.Y)/PlanarRegion.Scale*scale,0)));
                            }
                            loops.Add(CurveLoop.Create(curves));
                        }
                        var local=Extrude(loops,(top-bottom)*scale);
                        var transform=Transform.Identity;transform.BasisX=XYZ.BasisX/scale;transform.BasisY=XYZ.BasisY/scale;transform.BasisZ=XYZ.BasisZ/scale;transform.Origin=new XYZ(x,y,bottom);
                        var candidate=SolidUtils.CreateTransformed(local,transform);
                        if(candidate==null||!FinitePositiveVolume(candidate.Volume)||Math.Abs(candidate.Volume-expected)>Math.Max(1e-7,expected*1e-7))
                            throw new InvalidOperationException("Revit изменил площадь / объём при построении плоского результата.");
                        shape=candidate;break;
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){failures.Add("Масштаб "+scale.ToString(CultureInfo.InvariantCulture)+": "+ex.Message);}
                }
                if(shape==null)throw new InvalidOperationException("Плоский результат вычислен, но Revit не построил тело. Малые участки не удалены. "+string.Join(" | ",failures));
                if(planarBodies!=null)
                {var body=new LayeredBody{Curved=curved};body.Layers.Add(new PlanarLayer{Bottom=bottom,Top=top,Region=region});planarBodies[shape]=Tuple.Create(body,(string)null);}
                return shape;
            }
            private static bool FinitePositiveVolume(double value){return value>0&&!double.IsNaN(value)&&!double.IsInfinity(value);}
            private static List<Solid> BuildLayeredSolids(LayeredBody body)
            {
                var result=new List<Solid>();
                foreach(var layer in body.Layers)
                    foreach(var paths in layer.Region.Components())
                    {
                        planarCheckpoint?.Invoke();var region=new PlanarRegion(paths);
                        try{result.Add(BuildPlanarSolid(region,layer.Bottom,layer.Top,body.Curved));}
                        catch(System.OperationCanceledException){throw;}
                        catch(Exception original)
                        {
                            // Touching hole/outer boundaries can be valid polygons yet invalid
                            // Revit extrusion profiles. Split at every X vertex into simple strips.
                            var xs=paths.SelectMany(p=>p).Select(p=>p.X).Distinct().OrderBy(v=>v).ToList();
                            long minY=paths.SelectMany(p=>p).Min(p=>p.Y),maxY=paths.SelectMany(p=>p).Max(p=>p.Y);
                            var strips=new List<Solid>();double total=0;
                            try
                            {
                                for(int i=0;i+1<xs.Count;i++)
                                {
                                    planarCheckpoint?.Invoke();var box=new PlanarRegion(new List<List<CP>>{new List<CP>{new CP(xs[i],minY),new CP(xs[i+1],minY),new CP(xs[i+1],maxY),new CP(xs[i],maxY)}});
                                    var clipped=PlanarRegion.Combine(region,box,CT.ctIntersection);total+=clipped.Area;
                                    foreach(var part in clipped.Components())strips.Add(BuildPlanarSolid(new PlanarRegion(part),layer.Bottom,layer.Top,body.Curved));
                                }
                                if(Math.Abs(total-region.Area)>Math.Max(1e-8,region.Perimeter*4/PlanarRegion.Scale))throw new InvalidOperationException("Разбиение на полосы не сохранило площадь.");
                                if(strips.Count==0)throw new InvalidOperationException("Разбиение на полосы вернуло пустой результат.");
                                result.AddRange(strips);
                            }
                            catch(System.OperationCanceledException){throw;}
                            catch(Exception ex){throw new InvalidOperationException(original.Message+" | Разбиение плоского результата: "+ex.Message,ex);}
                        }
                    }
                return result;
            }
            private static Solid FlatBoolean(Solid a,Solid b,CT operation)
            {
                var aa=RequireLayeredBody(a);var bb=RequireLayeredBody(b);
                if(aa.Layers.Count!=1||bb.Layers.Count!=1)throw new InvalidOperationException("Операция площади требует одного плоского слоя у каждого контура.");
                var x=aa.Layers[0];var y=bb.Layers[0];
                if(Math.Abs(x.Bottom-y.Bottom)>1e-7||Math.Abs(x.Top-y.Top)>1e-7)throw new InvalidOperationException("Контуры площади лежат на разных расчётных плоскостях.");
                planarCheckpoint?.Invoke();
                var region=PlanarRegion.Combine(x.Region,y.Region,operation);
                if(operation==CT.ctDifference)
                {
                    var overlap=PlanarRegion.Combine(x.Region,y.Region,CT.ctIntersection);
                    double tolerance=Math.Max(1e-8,(x.Region.Perimeter+y.Region.Perimeter)*4/PlanarRegion.Scale);
                    if(Math.Abs(x.Region.Area-region.Area-overlap.Area)>tolerance)throw new InvalidOperationException("Нарушен баланс площадей плоского вычитания.");
                }
                return BuildPlanarSolid(region,x.Bottom,x.Top,aa.Curved||bb.Curved);
            }
            private bool TryLayeredDifference(Solid a,Solid b,out List<Solid> result,out string reason)
            {
                result=null;reason=null;LayeredBody aa,bb;string aReason,bReason;
                bool aOk=TryLayeredBody(a,out aa,out aReason),bOk=TryLayeredBody(b,out bb,out bReason);
                if(!aOk||!bOk){reason="A: "+(aReason??"вертикальные участки")+"; B: "+(bReason??"вертикальные участки");return false;}
                try
                {
                    var difference=LayeredBody.Combine(aa,bb,CT.ctDifference);
                    var overlap=LayeredBody.Combine(aa,bb,CT.ctIntersection);
                    double tolerance=Math.Max(1e-7,(aa.Layers.Concat(bb.Layers).Sum(l=>l.Region.Perimeter*(l.Top-l.Bottom)))*4/PlanarRegion.Scale);
                    if(Math.Abs(aa.Volume-difference.Volume-overlap.Volume)>tolerance)throw new InvalidOperationException("Нарушен баланс объёмов высотных участков.");
                    result=BuildLayeredSolids(difference);planarVolumeCuts++;return true;
                }
                catch(System.OperationCanceledException){throw;}
                catch(Exception ex){reason=ex.Message;return false;}
            }
            private static List<Solid> VolumeIntersections(Solid a,Solid b)
            {
                if(a==null||b==null||!Overlaps(Bounds(a),Bounds(b)))return new List<Solid>();
                LayeredBody aa,bb;string reason;
                if(TryLayeredBody(a,out aa,out reason)&&TryLayeredBody(b,out bb,out reason))
                    return BuildLayeredSolids(LayeredBody.Combine(aa,bb,CT.ctIntersection));
                var intersection=BooleanOperationsUtils.ExecuteBooleanOperation(a,b,BooleanOperationsType.Intersect);
                return intersection==null||intersection.Volume<1e-9?new List<Solid>():new List<Solid>{intersection};
            }
            private List<Solid> DecomposeVolumeInput(Solid solid,Record record,string metric)
            {
                LayeredBody body;string reason;
                if(TryLayeredBody(solid,out body,out reason)&&body.Layers.Count>1)
                {
                    try{return BuildLayeredSolids(body);}
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){Notice("PLANAR_REBUILD","Предупреждение","Не удалось представить исходное тело отдельными вертикальными участками; используется исходная геометрия. "+ex.Message,record,metric);}
                }
                return new List<Solid>{solid};
            }
            private static string BooleanObject(Record record)
            {
                return record==null?"не определён":record.Source.Name+"; ID "+IDHelper.ElIdValue(record.Element.Id)+"; UniqueId "+record.Element.UniqueId+
                    "; корпус "+record.Building+"; этаж "+(record.Level?.Name??"не определён");
            }

            // XY index for planar areas; XYZ index for volumes. Bounds are only a broad-phase filter.
            private sealed class PlanIndex
            {
                private readonly bool spatial;
                private readonly Dictionary<Tuple<int,int,int>,List<int>> cells=new Dictionary<Tuple<int,int,int>,List<int>>();
                private readonly List<int> large=new List<int>();
                private readonly List<Tuple<Solid,double[]>> entries=new List<Tuple<Solid,double[]>>();
                internal PlanIndex(bool spatial=false){this.spatial=spatial;}
                private static int Cell(double value){return checked((int)Math.Floor(value/32.0));}
                private int[] Range(double[] b){return new[]{Cell(b[0]),Cell(b[3]),Cell(b[1]),Cell(b[4]),spatial?Cell(b[2]):0,spatial?Cell(b[5]):0};}
                private static bool Large(int[] r){return ((double)r[1]-r[0]+1)*((double)r[3]-r[2]+1)*((double)r[5]-r[4]+1)>4096;}
                private readonly Dictionary<Solid,Record> owners=new Dictionary<Solid,Record>();
                internal Record OwnerOf(Solid solid){Record owner;return owners.TryGetValue(solid,out owner)?owner:null;}
                internal void Add(Solid solid,Record owner=null)
                {
                    if(solid==null||solid.Volume<1e-9)return;
                    if(owner!=null)owners[solid]=owner;
                    var b=Bounds(solid);var r=Range(b);int id=entries.Count;entries.Add(Tuple.Create(solid,b));
                    if(Large(r)){large.Add(id);return;}
                    for(int x=r[0];x<=r[1];x++)for(int y=r[2];y<=r[3];y++)for(int z=r[4];z<=r[5];z++)
                    {var key=Tuple.Create(x,y,z);List<int> bucket;if(!cells.TryGetValue(key,out bucket))cells[key]=bucket=new List<int>();bucket.Add(id);}
                }
                internal IEnumerable<Solid> Query(Solid solid)
                {
                    if(solid==null||solid.Volume<1e-9)yield break;
                    var b=Bounds(solid);var r=Range(b);var found=new HashSet<int>(large);
                    if(Large(r)){for(int i=0;i<entries.Count;i++)found.Add(i);}
                    else for(int x=r[0];x<=r[1];x++)for(int y=r[2];y<=r[3];y++)for(int z=r[4];z<=r[5];z++)
                    {List<int> bucket;if(cells.TryGetValue(Tuple.Create(x,y,z),out bucket))foreach(int id in bucket)found.Add(id);}
                    foreach(int id in found.OrderBy(i=>i))if(Overlaps(b,entries[id].Item2))yield return entries[id].Item1;
                }
            }
            private static double[] Bounds(Solid solid)
            {
                var box=solid.GetBoundingBox();var result=new[]{double.PositiveInfinity,double.PositiveInfinity,double.PositiveInfinity,double.NegativeInfinity,double.NegativeInfinity,double.NegativeInfinity};
                for(int i=0;i<8;i++)
                {var p=box.Transform.OfPoint(new XYZ((i&1)==0?box.Min.X:box.Max.X,(i&2)==0?box.Min.Y:box.Max.Y,(i&4)==0?box.Min.Z:box.Max.Z));
                    result[0]=Math.Min(result[0],p.X);result[1]=Math.Min(result[1],p.Y);result[2]=Math.Min(result[2],p.Z);
                    result[3]=Math.Max(result[3],p.X);result[4]=Math.Max(result[4],p.Y);result[5]=Math.Max(result[5],p.Z);}
                return result;
            }
            private static bool Overlaps(double[] a,double[] b)
            {return a[0]<b[3]&&a[3]>b[0]&&a[1]<b[4]&&a[4]>b[1]&&a[2]<b[5]&&a[5]>b[2];}
            private static Solid Union(Solid a,Solid b)
            {if(a==null||a.Volume<1e-9)return b;if(b==null||b.Volume<1e-9)return a;return FlatBoolean(a,b,CT.ctUnion);}
            private static Solid Subtract(Solid a,Solid b)
            {if(a==null||b==null||!Overlaps(Bounds(a),Bounds(b)))return a;return FlatBoolean(a,b,CT.ctDifference);}
            private static Solid Intersect(Solid a,Solid b)
            {if(a==null||b==null||!Overlaps(Bounds(a),Bounds(b)))return null;return FlatBoolean(a,b,CT.ctIntersection);}
            private static Curve Flat(Curve c,double z=0)
            {
                if(c is Line){var p=c.GetEndPoint(0);var q=c.GetEndPoint(1);return Line.CreateBound(new XYZ(p.X,p.Y,z),new XYZ(q.X,q.Y,z));}
                if(c is Arc)
                {var a=(Arc)c;if(Math.Abs(a.Normal.Z)<.999999)throw new InvalidOperationException("Наклонная дуга требует явного горизонтального расчётного контура.");
                    return c.CreateTransformed(Transform.CreateTranslation(new XYZ(0,0,z-c.GetEndPoint(0).Z)));}
                if(Math.Abs(c.GetEndPoint(0).Z-c.GetEndPoint(1).Z)>1e-7)throw new InvalidOperationException("Неплоская кривая расчётного контура.");
                return c.CreateTransformed(Transform.CreateTranslation(new XYZ(0,0,z-c.GetEndPoint(0).Z)));
            }
            private static CurveLoop FlatLoop(IEnumerable<Curve> curves,double z=0){return CurveLoop.Create(curves.Select(c=>Flat(c,z)).ToList());}
            private static Solid Extrude(IList<CurveLoop> loops,double height=1)
            {return GeometryCreationUtilities.CreateExtrusionGeometry(loops,XYZ.BasisZ,height);}
            private static IList<CurveLoop> BottomLoops(Solid solid,double z=0)
            {
                var faces=solid.Faces.Cast<Face>().OfType<PlanarFace>().Where(f=>f.FaceNormal.Z<-.999999).ToList();
                if(faces.Count!=1)throw new InvalidOperationException("Для обводки требуется один связный плоский контур; геометрия разбивается по граням.");
                return faces[0].GetEdgesAsCurveLoops().Select(l=>FlatLoop(l,z)).ToList();
            }
            private static Solid Projection(Solid solid)
            {
                LayeredBody body;string reason;if(TryLayeredBody(solid,out body,out reason))return BuildPlanarSolid(body.Projection(),0,1,body.Curved);
                Solid result=null;
                foreach(Face f in solid.Faces)
                {
                    var p=f as PlanarFace;if(p==null)
                    {var cylinder=f as CylindricalFace;if(cylinder!=null&&Math.Abs(cylinder.Axis.Z)>.999999)continue;throw new InvalidOperationException("Проекция криволинейной поверхности требует явной зоны / семейства расчётного контура.");}
                    if(p.FaceNormal.Z<=1e-7)continue;
                    result=Union(result,Extrude(p.GetEdgesAsCurveLoops().Select(l=>FlatLoop(l)).ToList()));
                }
                return result??throw new InvalidOperationException("Не найден замкнутый контур горизонтальной проекции.");
            }
            private static Solid Section(Solid solid,double z)
            {
                if(solid==null)return null;
                LayeredBody body;string reason;
                if(TryLayeredBody(solid,out body,out reason))
                {var profile=PlanarRegion.Empty;foreach(var l in body.Layers.Where(l=>l.Bottom<=z+1e-8&&l.Top>=z-1e-8))profile=PlanarRegion.Combine(profile,l.Region,CT.ctUnion);return BuildPlanarSolid(profile,0,1,body.Curved);}
                var box=solid.GetBoundingBox();
                var corners=Enumerable.Range(0,8).Select(i=>box.Transform.OfPoint(new XYZ((i&1)==0?box.Min.X:box.Max.X,(i&2)==0?box.Min.Y:box.Max.Y,(i&4)==0?box.Min.Z:box.Max.Z))).ToList();
                if(z<corners.Min(p=>p.Z)-1e-8||z>corners.Max(p=>p.Z)+1e-8)return null;
                Solid atBoundary=null;
                foreach(var face in solid.Faces.Cast<Face>().OfType<PlanarFace>().Where(f=>Math.Abs(Math.Abs(f.FaceNormal.Z)-1)<1e-10&&Math.Abs(f.Origin.Z-z)<1e-7))
                    atBoundary=Union(atBoundary,PlanFromSectionLoops(face.GetEdgesAsCurveLoops()));
                if(corners.Min(p=>p.Z)>=z||corners.Max(p=>p.Z)<=z)return atBoundary;
                // Translate to a local origin to improve native kernel conditioning. This does not
                // move the measurement plane relative to the solid or modify the source model.
                var origin=new XYZ((corners.Min(p=>p.X)+corners.Max(p=>p.X))/2,(corners.Min(p=>p.Y)+corners.Max(p=>p.Y))/2,z);
                var local=SolidUtils.CreateTransformed(solid,Transform.CreateTranslation(-origin));
                var failures=new List<string>();
                foreach(bool above in new[]{false,true})
                {
                    try
                    {
                        planarCheckpoint?.Invoke();
                        var cut=BooleanOperationsUtils.CutWithHalfSpace(local,Plane.CreateByNormalAndOrigin(above?XYZ.BasisZ:-XYZ.BasisZ,XYZ.Zero));
                        Solid result=atBoundary;bool found=false;
                        if(cut!=null)foreach(var face in cut.Faces.Cast<Face>().OfType<PlanarFace>().Where(f=>Math.Abs(Math.Abs(f.FaceNormal.Z)-1)<1e-10&&Math.Abs(f.Origin.Z)<1e-7))
                        {
                            var loops=face.GetEdgesAsCurveLoops().Select(l=>CurveLoop.Create(l.Select(c=>c.CreateTransformed(Transform.CreateTranslation(origin))).ToList())).ToList();
                            result=Union(result,PlanFromSectionLoops(loops));found=true;
                        }
                        if(found)return result;
                        failures.Add("Не получена грань точного сечения.");
                    }
                    catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){failures.Add(ex.Message);}
                }
                throw new InvalidOperationException("Revit не построил сечение на заданной отметке; высота не сдвигалась. "+string.Join(" | ",failures));
            }
            private static Solid Half(Solid solid,double z,bool above)
            {
                if(solid==null||solid.Volume<1e-9)return null;
                var box=solid.GetBoundingBox();var elevations=Enumerable.Range(0,8).Select(i=>box.Transform.OfPoint(new XYZ((i&1)==0?box.Min.X:box.Max.X,(i&2)==0?box.Min.Y:box.Max.Y,(i&4)==0?box.Min.Z:box.Max.Z)).Z).ToList();
                if(above){if(elevations.Max()<=z)return null;if(elevations.Min()>=z)return solid;}
                else{if(elevations.Min()>=z)return null;if(elevations.Max()<=z)return solid;}
                LayeredBody body;string reason;
                if(TryLayeredBody(solid,out body,out reason)&&body.Layers.Count==1)
                {var l=body.Layers[0];return BuildPlanarSolid(l.Region,above?Math.Max(z,l.Bottom):l.Bottom,above?l.Top:Math.Min(z,l.Top),body.Curved);}
                return BooleanOperationsUtils.CutWithHalfSpace(solid,Plane.CreateByNormalAndOrigin(above?XYZ.BasisZ:-XYZ.BasisZ,new XYZ(0,0,z)));
            }
            private static bool Inside(Solid solid,XYZ point)
            {
                using(var hit=solid.IntersectWithCurve(Line.CreateBound(new XYZ(point.X,point.Y,.25),new XYZ(point.X,point.Y,.75)),new SolidCurveIntersectionOptions()))return hit.SegmentCount>0;
            }
        }
    }
}
