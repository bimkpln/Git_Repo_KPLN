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
