using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        // Homogeneous rational Bezier subdivision. Positive weights give a convex-hull
        // bound, unlike midpoint-only sampling which can miss an entire S-shaped bend.
        internal static class RationalContours
        {
            internal static double[] Cartesian(double[] h)
            {
                if (h.Length != 4 || !(h[3] > 0) || h.Any(x => double.IsNaN(x) || double.IsInfinity(x)))
                    throw new InvalidOperationException("Некорректные контрольные точки / веса кривой.");
                return new[] { h[0] / h[3], h[1] / h[3], h[2] / h[3] };
            }
            internal static double DistanceToSegment(double[] p, double[] a, double[] b)
            {
                double dx=b[0]-a[0],dy=b[1]-a[1],dz=b[2]-a[2];
                double length=dx*dx+dy*dy+dz*dz;
                double t=length==0?0:Math.Max(0,Math.Min(1,((p[0]-a[0])*dx+(p[1]-a[1])*dy+(p[2]-a[2])*dz)/length));
                dx=p[0]-a[0]-t*dx;dy=p[1]-a[1]-t*dy;dz=p[2]-a[2]-t*dz;
                return Math.Sqrt(dx*dx+dy*dy+dz*dz);
            }
            internal static void Split(IList<double[]> points, out List<double[]> left, out List<double[]> right)
            {
                var work=points.Select(p=>(double[])p.Clone()).ToList();
                left=new List<double[]>{work[0]};right=new List<double[]>{work.Last()};
                while(work.Count>1)
                {
                    work=Enumerable.Range(0,work.Count-1).Select(i=>Enumerable.Range(0,4).Select(j=>(work[i][j]+work[i+1][j])/2).ToArray()).ToList();
                    left.Add(work[0]);right.Add(work.Last());
                }
                right.Reverse();
            }
            internal static List<double[]> Flatten(IList<double[]> points,double tolerance,Action checkpoint=null)
            {
                if(points.Count<2||!(tolerance>0))throw new ArgumentException("Некорректная кривая или допуск.");
                foreach(var p in points)Cartesian(p);
                var result=new List<double[]>();
                FlattenPart(points,tolerance,0,result,checkpoint);
                result.Add(Cartesian(points.Last()));
                return result;
            }
            private static void FlattenPart(IList<double[]> points,double tolerance,int depth,List<double[]> result,Action checkpoint)
            {
                checkpoint?.Invoke();
                if(result.Count>=200000||depth>40)throw new InvalidOperationException("Не удалось подтвердить точность кривой в пределах лимита разбиения. Допуск не увеличен.");
                var first=Cartesian(points[0]);var last=Cartesian(points.Last());
                if(points.All(p=>DistanceToSegment(Cartesian(p),first,last)<=tolerance)) {result.Add(first);return;}
                List<double[]> left,right;Split(points,out left,out right);
                FlattenPart(left,tolerance,depth+1,result,checkpoint);FlattenPart(right,tolerance,depth+1,result,checkpoint);
            }
            // One knot insertion in homogeneous coordinates (Piegl/Tiller A5.1, r=1).
            private static void Insert(int degree,List<double> knots,List<double[]> points,double value)
            {
                int n=points.Count-1,k=degree;
                while(k<n&&knots[k+1]<=value)k++;
                if(value==knots[n+1])k=n;
                int multiplicity=knots.Count(x=>x==value);
                var next=new double[points.Count+1][];
                for(int i=0;i<=k-degree;i++)next[i]=points[i];
                for(int i=k-multiplicity;i<=n;i++)next[i+1]=points[i];
                for(int i=k-degree+1;i<=k-multiplicity;i++)
                {
                    double denominator=knots[i+degree]-knots[i];
                    if(!(denominator>0))throw new InvalidOperationException("Вырожденный узловой вектор NURBS.");
                    double alpha=(value-knots[i])/denominator;
                    next[i]=Enumerable.Range(0,4).Select(j=>alpha*points[i][j]+(1-alpha)*points[i-1][j]).ToArray();
                }
                if(next.Any(p=>p==null))throw new InvalidOperationException("Не удалось разделить NURBS без потери контрольных точек.");
                knots.Insert(k+1,value);points.Clear();points.AddRange(next);
            }
            internal sealed class Span
            {internal double Start,End;internal List<double[]> Points;}
            internal static List<Span> BezierSpans(int degree,IEnumerable<double> sourceKnots,IEnumerable<double[]> sourcePoints,double start,double end)
            {
                var knots=sourceKnots.ToList();var points=sourcePoints.Select(p=>(double[])p.Clone()).ToList();
                if(degree<1||points.Count<=degree||knots.Count!=points.Count+degree+1||!(end>start)||
                    start<knots[degree]||end>knots[points.Count]||knots.Zip(knots.Skip(1),(a,b)=>a>b).Any(x=>x))
                    throw new InvalidOperationException("Некорректный параметрический диапазон NURBS.");
                foreach(var p in points)Cartesian(p);
                var values=knots.Where(x=>x>=start&&x<=end).Concat(new[]{start,end}).Distinct().OrderBy(x=>x).ToList();
                foreach(double value in values)
                    while(knots.Count(x=>x==value)<degree)Insert(degree,knots,points,value);
                var result=new List<Span>();
                for(int i=degree;i<points.Count;i++)
                    if(knots[i]>=start&&knots[i+1]<=end&&knots[i+1]>knots[i])
                        result.Add(new Span{Start=knots[i],End=knots[i+1],Points=points.Skip(i-degree).Take(degree+1).ToList()});
                if(result.Count==0)throw new InvalidOperationException("NURBS не содержит участков в диапазоне обмера.");
                return result;
            }
            internal static double[] Evaluate(IList<double[]> points,double parameter)
            {
                var work=points.Select(p=>(double[])p.Clone()).ToList();
                while(work.Count>1)work=Enumerable.Range(0,work.Count-1).Select(i=>Enumerable.Range(0,4).Select(j=>(1-parameter)*work[i][j]+parameter*work[i+1][j]).ToArray()).ToList();
                return Cartesian(work[0]);
            }
        }
    }
}
