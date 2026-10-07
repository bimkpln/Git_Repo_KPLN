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
        // Immutable numeric profiles. XY is in Revit feet; integer grid is 0.001 mm.
        internal sealed class PlanarRegion
        {
            internal const double Scale=304800.0;
            internal readonly List<List<CP>> Paths;
            internal static readonly PlanarRegion Empty=new PlanarRegion(new List<List<CP>>());
            internal PlanarRegion(List<List<CP>> paths){Paths=paths;}
            internal bool IsEmpty {get{return Paths.Count==0;}}
            internal static long Coordinate(double feet)
            {
                double value=feet*Scale;
                if(double.IsNaN(value)||double.IsInfinity(value)||Math.Abs(value)>1e13)throw new InvalidOperationException("Координата контура вне безопасного диапазона плоского расчёта.");
                return checked((long)Math.Round(value,MidpointRounding.AwayFromZero));
            }
            private static double RingArea(List<CP> path)
            {
                if(path.Count<3)return 0;double sum=0;var origin=path[0];
                for(int i=1;i+1<path.Count;i++)
                    sum+=((double)path[i].X-origin.X)*((double)path[i+1].Y-origin.Y)-((double)path[i+1].X-origin.X)*((double)path[i].Y-origin.Y);
                return sum/(2*Scale*Scale);
            }
            internal double Area {get{return Math.Abs(Paths.Sum(RingArea));}}
            internal double Perimeter
            {get{double sum=0;foreach(var path in Paths)for(int i=0;i<path.Count;i++){var p=path[i];var q=path[(i+1)%path.Count];double x=(double)p.X-q.X,y=(double)p.Y-q.Y;sum+=Math.Sqrt(x*x+y*y)/Scale;}return sum;}}
            internal static PlanarRegion FromRings(IEnumerable<IEnumerable<double[]>> rings)
            {
                var paths=new List<List<CP>>();
                foreach(var ring in rings)
                {
                    var path=new List<CP>();
                    foreach(var p in ring){var next=new CP(Coordinate(p[0]),Coordinate(p[1]));if(path.Count==0||path.Last()!=next)path.Add(next);}
                    if(path.Count>1&&path[0]==path.Last())path.RemoveAt(path.Count-1);
                    if(path.Count<3||RingArea(path)==0)throw new InvalidOperationException("Контур вырожден, самопересекается или меньше точности 0,001 мм. Он не был молча удалён.");
                    paths.Add(path);
                }
                return Execute(paths,new List<List<CP>>(),CT.ctUnion,PF.pftEvenOdd);
            }
            private static PlanarRegion Execute(List<List<CP>> a,List<List<CP>> b,CT operation,PF fill)
            {
                if(a.Count==0&&b.Count==0)return Empty;
                var clipper=new PC{StrictlySimple=true,PreserveCollinear=false};
                clipper.AddPaths(a,PT.ptSubject,true);clipper.AddPaths(b,PT.ptClip,true);
                var result=new List<List<CP>>();
                if(!clipper.Execute(operation,result,fill,fill))throw new InvalidOperationException("Не удалось выполнить плоскую операцию над замкнутыми контурами.");
                return new PlanarRegion(result);
            }
            internal static PlanarRegion Combine(PlanarRegion a,PlanarRegion b,CT operation)
            {
                a=a??Empty;b=b??Empty;
                if(b.IsEmpty)return operation==CT.ctIntersection?Empty:a;
                if(a.IsEmpty)return operation==CT.ctUnion||operation==CT.ctXor?b:Empty;
                // Normalized profiles have positive outer contours and negative holes.
                return Execute(a.Paths,b.Paths,operation,PF.pftNonZero);
            }
            internal bool Contains(double x,double y)
            {
                var point=new CP(Coordinate(x),Coordinate(y));bool inside=false;
                foreach(var ring in Paths)
                {
                    int hit=PC.PointInPolygon(point,ring);
                    if(hit<0)return true; // Closed physical boundary belongs to the material.
                    if(hit>0)inside=!inside;
                }
                return inside;
            }
            // Partition each segment at ALL polygon intersections. Endpoint/midpoint sampling
            // alone would accept a separator spanning a doorway or another gap in the wall.
            internal bool CoversPolyline(IList<double[]> points,Action checkpoint=null)
            {
                if(points.Count<2||IsEmpty)return false;
                foreach(var point in points)if(!Contains(point[0],point[1]))return false;
                for(int n=0;n+1<points.Count;n++)
                {
                    checkpoint?.Invoke();
                    double ax=points[n][0]*Scale,ay=points[n][1]*Scale;
                    double dx=(points[n+1][0]-points[n][0])*Scale,dy=(points[n+1][1]-points[n][1])*Scale;
                    double length2=dx*dx+dy*dy;if(length2==0)continue;
                    var cuts=new List<double>{0,1};
                    foreach(var ring in Paths)for(int i=0;i<ring.Count;i++)
                    {
                        var p=ring[i];var q=ring[(i+1)%ring.Count];
                        double ex=(double)q.X-p.X,ey=(double)q.Y-p.Y,px=p.X-ax,py=p.Y-ay;
                        double denominator=dx*ey-dy*ex;
                        if(denominator!=0)
                        {
                            double t=(px*ey-py*ex)/denominator,u=(px*dy-py*dx)/denominator;
                            if(t>0&&t<1&&u>=0&&u<=1)cuts.Add(t);
                        }
                        else if(px*dy-py*dx==0)
                        {
                            double t=(px*dx+py*dy)/length2,v=((q.X-ax)*dx+(q.Y-ay)*dy)/length2;
                            if(t>0&&t<1)cuts.Add(t);if(v>0&&v<1)cuts.Add(v);
                        }
                    }
                    cuts=cuts.Distinct().OrderBy(t=>t).ToList();
                    for(int i=0;i+1<cuts.Count;i++)
                    {
                        double t=(cuts[i]+cuts[i+1])/2;
                        if(!Contains((ax+dx*t)/Scale,(ay+dy*t)/Scale))return false;
                    }
                }
                return true;
            }
            // Constant-thickness physical strip along a straight boundary. Used for the
            // inner side of columns and joined elements; never substitutes a nominal width.
            internal double UniformBoundaryInset(double[] a,double[] b)
            {
                if(!SupportsBoundaryPolyline(Empty,new[]{this},new[]{a,b},4/Scale))throw new InvalidOperationException("Граница не подтверждена физическим сечением.");
                double dx=b[0]-a[0],dy=b[1]-a[1],length=Math.Sqrt(dx*dx+dy*dy);
                if(length<=0)throw new InvalidOperationException("Нулевая длина границы.");
                double nx=dy/length,ny=-dx/length,tolerance=4/Scale;
                var distances=Paths.SelectMany(r=>r).Select(p=>(p.X/Scale-a[0])*nx+(p.Y/Scale-a[1])*ny).ToList();
                double min=distances.Min(),max=distances.Max(),offset;
                if(min>=-tolerance&&max>tolerance)offset=max;
                else if(max<=tolerance&&min<-tolerance)offset=min;
                else throw new InvalidOperationException("Материал расположен с обеих сторон границы; внутренняя сторона неоднозначна.");
                var aa=new[]{a[0]+nx*offset,a[1]+ny*offset};var bb=new[]{b[0]+nx*offset,b[1]+ny*offset};
                var strip=FromRings(new[]{new[]{a,b,bb,aa}});
                if(!SupportsBoundaryPolyline(Empty,new[]{this},new[]{aa,bb},4/Scale)||Combine(strip,this,CT.ctDifference).Area>Math.Max(1e-9,strip.Perimeter*tolerance))
                    throw new InvalidOperationException("Толщина физического сечения меняется вдоль границы. Постоянное смещение не применено.");
                return offset;
            }
            private List<double[]> BoundaryIntervals(double[] a,double[] b,double tolerance)
            {
                double dx=b[0]-a[0],dy=b[1]-a[1],length=Math.Sqrt(dx*dx+dy*dy);
                var intervals=new List<double[]>();if(length==0)return intervals;
                foreach(var ring in Paths)for(int i=0;i<ring.Count;i++)
                {
                    var p=ring[i];var q=ring[(i+1)%ring.Count];
                    double px=p.X/Scale-a[0],py=p.Y/Scale-a[1],qx=q.X/Scale-a[0],qy=q.Y/Scale-a[1];
                    if(Math.Abs(px*dy-py*dx)>tolerance*length||Math.Abs(qx*dy-qy*dx)>tolerance*length)continue;
                    double u=(px*dx+py*dy)/(length*length),v=(qx*dx+qy*dy)/(length*length);
                    double start=Math.Max(0,Math.Min(u,v)),end=Math.Min(1,Math.Max(u,v));
                    // Only compensate endpoint rounding onto the integer grid; never extend
                    // internal intervals to bridge gaps between different elements.
                    if(start*length<=2/Scale)start=0;
                    if((1-end)*length<=2/Scale)end=1;
                    if(end>start)intervals.Add(new[]{start,end});
                }
                return intervals;
            }
            internal bool HasBoundaryOverlap(IList<double[]> points,double tolerance)
            {
                for(int i=0;i+1<points.Count;i++)
                    if(BoundaryIntervals(points[i],points[i+1],tolerance).Any())return true;
                return false;
            }
            // A Revit boundary can contain the working frame and building loops joined by
            // opposite traversals of a zero-area seam. Remove only identified frame edges;
            // the ordinary graph validator cancels seams and still rejects actual gaps.
            internal static PlanarRegion FromExteriorBoundarySegments(IList<double[][]> segments,IList<bool> calculationFrame,Action checkpoint=null)
            {
                if(segments==null||calculationFrame==null||segments.Count!=calculationFrame.Count)
                    throw new ArgumentException("Не совпадают участки границы и признаки расчётной рамки.");
                return FromBoundarySegments(segments.Where((s,i)=>!calculationFrame[i]).ToList(),checkpoint);
            }
            internal static bool SupportsBoundaryPolyline(PlanarRegion walls,IEnumerable<PlanarRegion> otherFaces,IList<double[]> points,double tolerance,Action checkpoint=null)
            {
                if(points.Count<2)return false;
                var faces=otherFaces.ToList();
                for(int n=0;n+1<points.Count;n++)
                {
                    checkpoint?.Invoke();var a=points[n];var b=points[n+1];
                    double dx=b[0]-a[0],dy=b[1]-a[1];if(dx==0&&dy==0)continue;
                    var intervals=faces.SelectMany(r=>r.BoundaryIntervals(a,b,tolerance)).ToList();
                    var cuts=intervals.SelectMany(i=>i).Concat(new[]{0.0,1.0}).ToList();
                    foreach(var ring in walls.Paths)for(int i=0;i<ring.Count;i++)
                    {
                        var p=ring[i];var q=ring[(i+1)%ring.Count];
                        double px=p.X/Scale-a[0],py=p.Y/Scale-a[1],ex=((double)q.X-p.X)/Scale,ey=((double)q.Y-p.Y)/Scale;
                        double denominator=dx*ey-dy*ex;
                        if(denominator!=0)
                        {
                            double t=(px*ey-py*ex)/denominator,u=(px*dy-py*dx)/denominator;
                            if(t>0&&t<1&&u>=0&&u<=1)cuts.Add(t);
                        }
                        else if(px*dy-py*dx==0)
                        {
                            double length2=dx*dx+dy*dy;
                            double t=(px*dx+py*dy)/length2,v=((px+ex)*dx+(py+ey)*dy)/length2;
                            if(t>0&&t<1)cuts.Add(t);if(v>0&&v<1)cuts.Add(v);
                        }
                    }
                    cuts=cuts.Distinct().OrderBy(t=>t).ToList();
                    Func<double,bool> supported=t=>walls.Contains(a[0]+dx*t,a[1]+dy*t)||intervals.Any(i=>i[0]<=t&&i[1]>=t);
                    if(!supported(0)||!supported(1))return false;
                    for(int i=0;i+1<cuts.Count;i++)if(!supported((cuts[i]+cuts[i+1])/2))return false;
                }
                return true;
            }
            internal sealed class RoomCell
            {
                internal PlanarRegion Region;
                internal List<int> Owners;
            }
            internal static List<RoomCell> PartitionByRooms(PlanarRegion envelope,IList<PlanarRegion> rooms,Action checkpoint=null)
            {
                var cells=new List<RoomCell>();var covered=Empty;
                for(int i=0;i<rooms.Count;i++)
                {
                    checkpoint?.Invoke();var centre=Combine(envelope,rooms[i],CT.ctIntersection);
                    if(Combine(centre,covered,CT.ctIntersection).Area*.09290304>.005)
                        throw new InvalidOperationException("Сечения помещений по осям ограждений пересекаются более чем на 0,005 м²; индекс помещения "+i);
                    centre=Combine(centre,covered,CT.ctDifference);covered=Combine(covered,centre,CT.ctUnion);
                    cells.Add(new RoomCell{Region=centre,Owners=new List<int>{i}});
                }
                foreach(var paths in Combine(envelope,covered,CT.ctDifference).Components())
                {
                    checkpoint?.Invoke();var region=new PlanarRegion(paths);
                    var neighbours=cells.Where(c=>c.Owners.Count==1&&c.Region.SharesBoundary(region)).Select(c=>c.Owners[0]).Distinct().ToList();
                    if(neighbours.Count==0)throw new InvalidOperationException("Не определена принадлежность участка конструкций площадью "+(region.Area*.09290304).ToString("0.######",CultureInfo.InvariantCulture)+" м²: нет общей границы с помещениями. Ближайшее помещение не назначено автоматически.");
                    if(neighbours.Count==1)
                    {
                        var cell=cells.First(c=>c.Owners.Count==1&&c.Owners[0]==neighbours[0]);cell.Region=Combine(cell.Region,region,CT.ctUnion);
                    }
                    else cells.Add(new RoomCell{Region=region,Owners=neighbours});
                }
                if(Math.Abs(cells.Sum(c=>c.Region.Area)-envelope.Area)*.09290304>.001)throw new InvalidOperationException("Не сошёлся баланс разделения этажа на помещения и участки конструкций.");
                return cells;
            }
            internal static bool CommonCellInclusion(IEnumerable<bool> included)
            {
                var values=included.Distinct().ToList();
                if(values.Count!=1)throw new InvalidOperationException("У общего участка конструкций не определено единственное правило включения.");
                return values[0];
            }
            internal bool SharesBoundary(PlanarRegion other)
            {
                foreach(var ring in Paths)for(int i=0;i<ring.Count;i++)
                {
                    var a=ring[i];var b=ring[(i+1)%ring.Count];double dx=(double)b.X-a.X,dy=(double)b.Y-a.Y,length=Math.Sqrt(dx*dx+dy*dy);
                    if(length==0)continue;
                    foreach(var second in other.Paths)for(int j=0;j<second.Count;j++)
                    {
                        var p=second[j];var q=second[(j+1)%second.Count];
                        if(Math.Max(a.X,b.X)+2<Math.Min(p.X,q.X)||Math.Max(p.X,q.X)+2<Math.Min(a.X,b.X)||Math.Max(a.Y,b.Y)+2<Math.Min(p.Y,q.Y)||Math.Max(p.Y,q.Y)+2<Math.Min(a.Y,b.Y))continue;
                        double px=(double)p.X-a.X,py=(double)p.Y-a.Y,qx=(double)q.X-a.X,qy=(double)q.Y-a.Y;
                        if(Math.Abs(px*dy-py*dx)>2*length||Math.Abs(qx*dy-qy*dx)>2*length)continue;
                        double u=(px*dx+py*dy)/length,v=(qx*dx+qy*dy)/length;
                        if(Math.Min(length,Math.Max(u,v))-Math.Max(0,Math.Min(u,v))>2)return true;
                    }
                }
                return false;
            }
            // Assemble a planar graph from native segments, without creating a Revit CurveLoop.
            // Only real intersections and collinear overlaps are noded. No gap is closed.
            internal static PlanarRegion FromBoundarySegments(IList<double[][]> lines,Action checkpoint=null)
            {
                var edges=new List<Tuple<CP,CP,List<CP>>>();
                foreach(var line in lines)for(int i=0;i+1<line.Length;i++)
                {
                    var a=new CP(Coordinate(line[i][0]),Coordinate(line[i][1]));var b=new CP(Coordinate(line[i+1][0]),Coordinate(line[i+1][1]));
                    if(a!=b)edges.Add(Tuple.Create(a,b,new List<CP>{a,b}));
                }
                Func<CP,CP,CP,double> cross=(a,b,p)=>((double)b.X-a.X)*((double)p.Y-a.Y)-((double)b.Y-a.Y)*((double)p.X-a.X);
                Action<Tuple<CP,CP,List<CP>>,CP> add=(e,p)=>{
                    double dx=(double)e.Item2.X-e.Item1.X,dy=(double)e.Item2.Y-e.Item1.Y,length2=dx*dx+dy*dy;
                    double t=(((double)p.X-e.Item1.X)*dx+((double)p.Y-e.Item1.Y)*dy)/length2;
                    if(t>=0&&t<=1&&Math.Abs(cross(e.Item1,e.Item2,p))<=2*Math.Sqrt(length2))e.Item3.Add(p);
                };
                var active=new List<Tuple<CP,CP,List<CP>>>();
                foreach(var e in edges.OrderBy(e=>Math.Min(e.Item1.X,e.Item2.X)))
                {
                    checkpoint?.Invoke();long left=Math.Min(e.Item1.X,e.Item2.X);
                    active.RemoveAll(f=>Math.Max(f.Item1.X,f.Item2.X)<left-2);
                    foreach(var f in active)
                    {
                        if(Math.Min(e.Item1.Y,e.Item2.Y)>Math.Max(f.Item1.Y,f.Item2.Y)+2||Math.Max(e.Item1.Y,e.Item2.Y)<Math.Min(f.Item1.Y,f.Item2.Y)-2)continue;
                        add(e,f.Item1);add(e,f.Item2);add(f,e.Item1);add(f,e.Item2);
                        double dx=(double)e.Item2.X-e.Item1.X,dy=(double)e.Item2.Y-e.Item1.Y,ex=(double)f.Item2.X-f.Item1.X,ey=(double)f.Item2.Y-f.Item1.Y;
                        if(e.Item3.Intersect(f.Item3).Take(2).Count()==2)continue;
                        double lenE=Math.Sqrt(dx*dx+dy*dy),lenF=Math.Sqrt(ex*ex+ey*ey);
                        // Rounding collinear source lines onto the grid can give them tiny
                        // different slopes. Do not create a spurious crossing between them.
                        if(Math.Abs(cross(e.Item1,e.Item2,f.Item1))<=2*lenE&&Math.Abs(cross(e.Item1,e.Item2,f.Item2))<=2*lenE&&
                            Math.Abs(cross(f.Item1,f.Item2,e.Item1))<=2*lenF&&Math.Abs(cross(f.Item1,f.Item2,e.Item2))<=2*lenF)continue;
                        double den=dx*ey-dy*ex;if(den==0)continue;
                        double px=(double)f.Item1.X-e.Item1.X,py=(double)f.Item1.Y-e.Item1.Y;
                        double t=(px*ey-py*ex)/den,u=(px*dy-py*dx)/den;
                        if(t>0&&t<1&&u>0&&u<1)
                        {
                            var point=new CP((long)Math.Round(e.Item1.X+dx*t,MidpointRounding.AwayFromZero),(long)Math.Round(e.Item1.Y+dy*t,MidpointRounding.AwayFromZero));
                            var shared=e.Item3.Intersect(f.Item3).Where(p=>Math.Abs(p.X-point.X)<=2&&Math.Abs(p.Y-point.Y)<=2).ToList();
                            if(shared.Count==1)point=shared[0]; // same proven node, rounded intersection
                            e.Item3.Add(point);f.Item3.Add(point);
                        }
                    }
                    active.Add(e);
                }
                Func<CP,CP,int> compare=(a,b)=>a.X!=b.X?a.X.CompareTo(b.X):a.Y.CompareTo(b.Y);
                var balance=new Dictionary<Tuple<CP,CP>,int>();
                foreach(var e in edges)
                {
                    double dx=(double)e.Item2.X-e.Item1.X,dy=(double)e.Item2.Y-e.Item1.Y;
                    var points=e.Item3.Distinct().OrderBy(p=>((double)p.X-e.Item1.X)*dx+((double)p.Y-e.Item1.Y)*dy).ToList();
                    for(int i=0;i+1<points.Count;i++)
                    {
                        var a=points[i];var b=points[i+1];if(a==b)continue;bool forward=compare(a,b)<0;
                        var key=forward?Tuple.Create(a,b):Tuple.Create(b,a);int count;balance.TryGetValue(key,out count);balance[key]=count+(forward?1:-1);
                    }
                }
                var links=new Dictionary<CP,List<CP>>();
                foreach(var e in balance.Where(e=>e.Value!=0).Select(e=>e.Key))
                {
                    if(!links.ContainsKey(e.Item1))links[e.Item1]=new List<CP>();if(!links.ContainsKey(e.Item2))links[e.Item2]=new List<CP>();
                    links[e.Item1].Add(e.Item2);links[e.Item2].Add(e.Item1);
                }
                var invalid=links.Where(p=>p.Value.Count!=2).Take(3).ToList();
                if(invalid.Count>0)throw new InvalidOperationException("В графе границ есть разрыв или неоднозначное пересечение: "+string.Join("; ",invalid.Select(p=>"XY "+(p.Key.X/Scale*.3048).ToString("0.######",CultureInfo.InvariantCulture)+", "+(p.Key.Y/Scale*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м; ветвей "+p.Value.Count))+". Соединяющие отрезки не добавлены.");
                if(links.Count==0)throw new InvalidOperationException("Граница не содержит замкнутых областей после устранения обратных повторов.");
                var visited=new HashSet<CP>();var rings=new List<double[][]>();
                foreach(var start in links.Keys)
                {
                    if(visited.Contains(start))continue;var points=new List<double[]>();var current=start;CP previous=start;
                    do
                    {
                        if(!visited.Add(current))throw new InvalidOperationException("Неоднозначное ветвление замкнутой границы.");
                        points.Add(new[]{current.X/Scale,current.Y/Scale});var next=links[current].First(p=>p!=previous);previous=current;current=next;
                    }while(current!=start);
                    rings.Add(points.ToArray());
                }
                ValidateBoundaryRings(rings);
                return FromRings(rings);
            }

            internal PlanarRegion BoundaryBandsFromSegments(IList<double[][]> lines,IList<double?> offsets)
            {
                if(lines.Count!=offsets.Count)throw new ArgumentException("Не совпадают границы и смещения.");
                var result=Empty;
                foreach(var path in Paths)
                {
                    var points=new List<double[]>();var widths=new List<double>();
                    for(int edge=0;edge<path.Count;edge++)
                    {
                        double ax=path[edge].X/Scale,ay=path[edge].Y/Scale,bx=path[(edge+1)%path.Count].X/Scale,by=path[(edge+1)%path.Count].Y/Scale;
                        double dx=bx-ax,dy=by-ay,length=Math.Sqrt(dx*dx+dy*dy),length2=length*length;
                        var covers=new List<Tuple<double,double,double?>>();var cuts=new List<double>{0,1};
                        for(int i=0;i<lines.Count;i++)
                        {
                            var a=lines[i][0];var b=lines[i][1];
                            if(Math.Abs((a[0]-ax)*dy-(a[1]-ay)*dx)>2/Scale*length||Math.Abs((b[0]-ax)*dy-(b[1]-ay)*dx)>2/Scale*length)continue;
                            double u=((a[0]-ax)*dx+(a[1]-ay)*dy)/length2,v=((b[0]-ax)*dx+(b[1]-ay)*dy)/length2;
                            double low=Math.Max(0,Math.Min(u,v)),high=Math.Min(1,Math.Max(u,v));
                            if((high-low)*length<=1/Scale)continue;
                            if(low*length<=2/Scale)low=0;if((1-high)*length<=2/Scale)high=1;
                            covers.Add(Tuple.Create(low,high,offsets[i].HasValue?(double?)(offsets[i].Value*(v>=u?1:-1)):null));cuts.Add(low);cuts.Add(high);
                        }
                        cuts=cuts.Distinct().OrderBy(t=>t).ToList();
                        for(int i=0;i+1<cuts.Count;i++)
                        {
                            if((cuts[i+1]-cuts[i])*length<=1/Scale)continue;
                            double mid=(cuts[i]+cuts[i+1])/2;
                            var values=covers.Where(c=>c.Item1<=mid&&c.Item2>=mid&&c.Item3.HasValue).Select(c=>c.Item3.Value).ToList();
                            if(values.Count==0)throw new InvalidOperationException("Не определена внутренняя сторона действующего участка границы XY "+((ax+dx*mid)*.3048).ToString("0.######",CultureInfo.InvariantCulture)+", "+((ay+dy*mid)*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м. Неизвестная толщина не заменена средней.");
                            if(values.Max()-values.Min()>4/Scale)throw new InvalidOperationException("Разные ограждения задают противоречащие внутренние стороны одного участка.");
                            points.Add(new[]{ax+dx*cuts[i],ay+dy*cuts[i]});widths.Add(values[0]);
                        }
                    }
                    result=Combine(result,StraightBoundaryBands(points.ToArray(),widths.ToArray()),CT.ctUnion);
                }
                return result;
            }

            internal static PlanarRegion StraightBoundaryBands(double[][] ring,double[] offsets)
            {
                if(ring.Length<3||ring.Length!=offsets.Length)throw new ArgumentException("Неполная граница полос ограждения.");
                ValidateBoundaryRings(new[]{ring});
                var starts=new List<double[]>();var ends=new List<double[]>();var directions=new List<double[]>();
                var result=Empty;
                for(int i=0;i<ring.Length;i++)
                {
                    var a=ring[i];var b=ring[(i+1)%ring.Length];double dx=b[0]-a[0],dy=b[1]-a[1],length=Math.Sqrt(dx*dx+dy*dy);
                    if(length==0)throw new InvalidOperationException("Нулевая длина участка ограждения.");
                    double nx=dy/length*offsets[i],ny=-dx/length*offsets[i];
                    var aa=new[]{a[0]+nx,a[1]+ny};var bb=new[]{b[0]+nx,b[1]+ny};
                    starts.Add(aa);ends.Add(bb);directions.Add(new[]{dx,dy});
                    if(Math.Abs(offsets[i])>1e-12)result=Combine(result,FromRings(new[]{new[]{a,b,bb,aa}}),CT.ctUnion);
                }
                for(int i=0;i<ring.Length;i++)
                {
                    int next=(i+1)%ring.Length;
                    result=Combine(result,BoundaryBandJoin(ring[next],ends[i],directions[i],starts[next],directions[next]),CT.ctUnion);
                }
                return result;
            }
            // Fill the join between two independently measured straight boundary bands.
            // No curve-loop trimming or tolerance-based closing of a model gap is performed.
            internal static PlanarRegion BoundaryBandJoin(double[] corner,double[] a,double[] da,double[] b,double[] db)
            {
                double dx=b[0]-a[0],dy=b[1]-a[1];
                if(Math.Sqrt(dx*dx+dy*dy)<=2/Scale)return Empty;
                double denominator=da[0]*db[1]-da[1]*db[0];
                double lengthA=Math.Sqrt(da[0]*da[0]+da[1]*da[1]),lengthB=Math.Sqrt(db[0]*db[0]+db[1]*db[1]);
                if(lengthA==0||lengthB==0)throw new InvalidOperationException("Нулевая касательная внутренней границы.");
                if(Math.Abs(denominator)<=1e-12*lengthA*lengthB)
                    throw new InvalidOperationException("Параллельные участки имеют разные внутренние границы; соединение не выдумано.");
                double t=(dx*db[1]-dy*db[0])/denominator;
                var m=new[]{a[0]+da[0]*t,a[1]+da[1]*t};
                var pieces=new[]{new[]{corner,a,m},new[]{corner,m,b}};
                var result=Empty;
                foreach(var triangle in pieces)
                {
                    double cross=(triangle[1][0]-triangle[0][0])*(triangle[2][1]-triangle[0][1])-(triangle[1][1]-triangle[0][1])*(triangle[2][0]-triangle[0][0]);
                    if(Math.Abs(cross)>1/(Scale*Scale))result=Combine(result,FromRings(new[]{triangle}),CT.ctUnion);
                }
                return result;
            }
            internal static void ValidateBoundaryRings(IList<double[][]> rings)
            {
                var segments=new List<Tuple<int,int,int,CP,CP>>();
                for(int ring=0;ring<rings.Count;ring++)
                {
                    var points=rings[ring].Select(p=>new CP(Coordinate(p[0]),Coordinate(p[1]))).ToList();
                    for(int i=points.Count-1;i>0;i--)if(points[i]==points[i-1])points.RemoveAt(i);
                    if(points.Count>1&&points[0]==points.Last())points.RemoveAt(points.Count-1);
                    if(points.Count<3)throw new InvalidOperationException("Вырожденная граница этажа.");
                    for(int i=0;i<points.Count;i++)segments.Add(Tuple.Create(ring,i,points.Count,points[i],points[(i+1)%points.Count]));
                }
                Func<CP,CP,CP,double> side=(a,b,p)=>((double)b.X-a.X)*((double)p.Y-a.Y)-((double)b.Y-a.Y)*((double)p.X-a.X);
                var active=new List<Tuple<int,int,int,CP,CP>>();
                foreach(var edge in segments.OrderBy(e=>Math.Min(e.Item4.X,e.Item5.X)))
                {
                    long left=Math.Min(edge.Item4.X,edge.Item5.X);
                    active.RemoveAll(e=>Math.Max(e.Item4.X,e.Item5.X)<left);
                    foreach(var other in active)
                    {
                        if(edge.Item1==other.Item1&&(Math.Abs(edge.Item2-other.Item2)==1||Math.Abs(edge.Item2-other.Item2)==edge.Item3-1))continue;
                        if(Math.Min(edge.Item4.Y,edge.Item5.Y)>Math.Max(other.Item4.Y,other.Item5.Y)||Math.Max(edge.Item4.Y,edge.Item5.Y)<Math.Min(other.Item4.Y,other.Item5.Y))continue;
                        double a=side(edge.Item4,edge.Item5,other.Item4),b=side(edge.Item4,edge.Item5,other.Item5);
                        double c=side(other.Item4,other.Item5,edge.Item4),d=side(other.Item4,other.Item5,edge.Item5);
                        if(((a<=0&&b>=0)||(a>=0&&b<=0))&&((c<=0&&d>=0)||(c>=0&&d<=0)))
                            throw new InvalidOperationException("Границы этажа пересекаются или касаются вне соседних вершин. Площадь не принята автоматически.");
                    }
                    active.Add(edge);
                }
            }
            internal IEnumerable<List<List<CP>>> Components()
            {
                var clipper=new PC{StrictlySimple=true};clipper.AddPaths(Paths,PT.ptSubject,true);
                var tree=new TepClipper.PolyTree();
                if(!IsEmpty&&!clipper.Execute(CT.ctUnion,tree,PF.pftNonZero,PF.pftNonZero))throw new InvalidOperationException("Не удалось выделить компоненты контура.");
                for(var node=tree.GetFirst();node!=null;node=node.GetNext())
                    if(!node.IsHole&&!node.IsOpen&&node.Contour.Count>0)
                    {var component=new List<List<CP>>{node.Contour};component.AddRange(node.Childs.Where(n=>n.IsHole).Select(n=>n.Contour));yield return component;}
            }
        }
        internal sealed class PlanarLayer
        {
            internal double Bottom,Top;
            internal PlanarRegion Region;
            internal double Volume {get{return Region.Area*(Top-Bottom);}}
        }
        internal sealed class LayeredBody
        {
            internal readonly List<PlanarLayer> Layers=new List<PlanarLayer>();
            internal bool Curved;
            internal double Volume {get{return Layers.Sum(l=>l.Volume);}}
            internal PlanarRegion At(double z)
            {return Layers.FirstOrDefault(l=>l.Bottom<=z&&l.Top>z)?.Region??PlanarRegion.Empty;}
            internal PlanarRegion Projection()
            {var result=PlanarRegion.Empty;foreach(var l in Layers)result=PlanarRegion.Combine(result,l.Region,CT.ctUnion);return result;}
            // Projection of the body between the lowest calculation floor and ground,
            // including the section at ground. A body ending at the lowest floor is below scope.
            internal PlanarRegion Footprint(double lowest, double ground)
            {
                if (lowest > ground) return PlanarRegion.Empty;
                var result = PlanarRegion.Empty;
                foreach (var layer in Layers.Where(l => l.Top > lowest && l.Bottom <= ground))
                    result = PlanarRegion.Combine(result, layer.Region, CT.ctUnion);
                return result;
            }
            internal static LayeredBody Combine(LayeredBody a,LayeredBody b,CT operation)
            {
                var levels=a.Layers.Concat(b.Layers).SelectMany(l=>new[]{l.Bottom,l.Top}).Distinct().OrderBy(z=>z).ToList();
                var result=new LayeredBody{Curved=a.Curved||b.Curved};
                for(int i=0;i+1<levels.Count;i++)
                {
                    double z=levels[i]+(levels[i+1]-levels[i])/2;
                    var area=PlanarRegion.Combine(a.At(z),b.At(z),operation);
                    if(!area.IsEmpty)result.Layers.Add(new PlanarLayer{Bottom=levels[i],Top=levels[i+1],Region=area});
                }
                return result;
            }
        }
    }
}
