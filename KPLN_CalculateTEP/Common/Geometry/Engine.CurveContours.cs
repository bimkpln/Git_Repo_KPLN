using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            public static double ConicRawParameter(double start,double end,double t)
            {
                if(!(end>start)||end-start>Math.PI/2+1e-12||t<0||t>1)throw new ArgumentOutOfRangeException();
                if(t==0)return start;if(t==1)return end;
                // Rational quadratic weight cancels the middle control point's 1/cos(half-angle).
                // Its Bezier parameter is NOT a linear angular parameter of the Revit arc/ellipse.
                double middle=(start+end)/2,a=(1-t)*(1-t),b=2*t*(1-t),c=t*t;
                double angle=Math.Atan2(a*Math.Sin(start)+b*Math.Sin(middle)+c*Math.Sin(end),
                    a*Math.Cos(start)+b*Math.Cos(middle)+c*Math.Cos(end));
                return angle+2*Math.PI*Math.Round((middle-angle)/(2*Math.PI));
            }
            private static List<double[]> CurvePoints(Curve curve, double chordTolerance = ArcChordTolerance)
            {
                if(curve is Line)return new[]{curve.GetEndPoint(0),curve.GetEndPoint(1)}.Select(p=>new[]{p.X,p.Y,p.Z}).ToList();
                var hermite=curve as HermiteSpline;
                var spans=new List<RationalContours.Span>();
                if(hermite!=null)
                {
                    // Cubic Hermite intervals convert directly to Bezier, including periodic curves.
                    // Use native derivatives with respect to the raw curve parameter, not unit tangents.
                    var knots=hermite.Parameters.Cast<double>().ToList();
                    double start=curve.IsBound?curve.GetEndParameter(0):knots.First();
                    double end=curve.IsBound?curve.GetEndParameter(1):start+curve.Period;
                    var breaks=knots.Where(t=>t>start&&t<end).Concat(new[]{start,end}).Distinct().OrderBy(t=>t).ToList();
                    for(int i=0;i+1<breaks.Count;i++)
                    {
                        double a=breaks[i],b=breaks[i+1],third=(b-a)/3;
                        var p=curve.Evaluate(a,false);var q=curve.Evaluate(b,false);
                        var controls=new[]{p,p+curve.ComputeDerivatives(a,false).BasisX*third,q-curve.ComputeDerivatives(b,false).BasisX*third,q};
                        spans.Add(new RationalContours.Span{Start=a,End=b,Points=controls.Select(x=>new[]{x.X,x.Y,x.Z,1.0}).ToList()});
                    }
                }
                var ellipse=curve as Ellipse;var arc=curve as Arc;
                if(hermite!=null) { /* Intervals prepared above; all intervals are validated below. */ }
                else if(ellipse!=null||arc!=null)
                {
                    XYZ center=arc!=null?arc.Center:ellipse.Center;
                    XYZ x=(arc!=null?arc.XDirection:ellipse.XDirection)*(arc!=null?arc.Radius:ellipse.RadiusX);
                    XYZ y=(arc!=null?arc.YDirection:ellipse.YDirection)*(arc!=null?arc.Radius:ellipse.RadiusY);
                    double a=curve.IsBound?curve.GetEndParameter(0):0,b=curve.IsBound?curve.GetEndParameter(1):2*Math.PI;
                    int count=Math.Max(1,(int)Math.Ceiling((b-a)/(Math.PI/2)));
                    for(int i=0;i<count;i++)
                    {
                        double u=a+(b-a)*i/count,v=a+(b-a)*(i+1)/count,m=(u+v)/2,w=Math.Cos((v-u)/2);
                        var p=center+x*Math.Cos(u)+y*Math.Sin(u);var q=center+(x*Math.Cos(m)+y*Math.Sin(m))/w;var r=center+x*Math.Cos(v)+y*Math.Sin(v);
                        spans.Add(new RationalContours.Span{Start=u,End=v,Points=new List<double[]>{new[]{p.X,p.Y,p.Z,1.0},new[]{q.X*w,q.Y*w,q.Z*w,w},new[]{r.X,r.Y,r.Z,1.0}}});
                    }
                }
                else
                {
                    var spline=curve as NurbSpline;
                    if(spline==null)throw new InvalidOperationException("Revit вернул кривую "+curve.GetType().Name+", для которой не предоставлено проверяемое представление NURBS. Нет подмены прямой или грубой Tessellate.");
                    var knots=spline.Knots.Cast<double>().ToList();var controls=spline.CtrlPoints;var weights=spline.Weights.Cast<double>().ToList();
                    if(weights.Count==0)weights=Enumerable.Repeat(1.0,controls.Count).ToList();
                    if(weights.Count!=controls.Count)throw new InvalidOperationException("Число весов NURBS не соответствует числу контрольных точек.");
                    var homogeneous=controls.Select((p,i)=>new[]{p.X*weights[i],p.Y*weights[i],p.Z*weights[i],weights[i]});
                    double start=curve.IsBound?curve.GetEndParameter(0):knots[spline.Degree],end=curve.IsBound?curve.GetEndParameter(1):knots[controls.Count];
                    spans=RationalContours.BezierSpans(spline.Degree,knots,homogeneous,start,end);
                }
                var points=new List<double[]>();
                foreach(var span in spans)
                {
                    foreach(double t in new[]{0.0,.25,.5,.75,1.0})
                    {
                        double raw=arc!=null||ellipse!=null?ConicRawParameter(span.Start,span.End,t):span.Start+(span.End-span.Start)*t;
                        var p=RationalContours.Evaluate(span.Points,t);var expected=curve.Evaluate(raw,false);
                        if(expected.DistanceTo(new XYZ(p[0],p[1],p[2]))>1e-7)throw new InvalidOperationException("Параметризация кривой Revit не совпала с рациональным представлением. Точность не подтверждена.");
                    }
                    var part=RationalContours.Flatten(span.Points,chordTolerance,planarCheckpoint);
                    if(points.Count>0)points.RemoveAt(points.Count-1);points.AddRange(part);
                    if(points.Count>200000)throw new InvalidOperationException("Слишком сложная кривая для установленного допуска 0,1 мм.");
                }
                if(points.Count<2||points.Max(p=>p[2])-points.Min(p=>p[2])>1e-7)throw new InvalidOperationException("Граница сечения не лежит в горизонтальной плоскости.");
                return points;
            }
            private static Solid PlanFromSectionLoops(IEnumerable<CurveLoop> loops, double chordTolerance = ArcChordTolerance)
            {
                bool curved=false;
                var region=PlanarRegion.FromRings(loops.Select(l=>PolygonRing(l,ref curved,chordTolerance)).ToList());
                return BuildPlanarSolid(region,0,1,curved);
            }
        }
    }
}
