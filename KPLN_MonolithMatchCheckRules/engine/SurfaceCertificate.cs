using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;

// Exact point/triangle distances in a bounding-volume hierarchy. Bounds for a
// source triangle use convexity of distance to ONE fixed target triangle and
// the 1-Lipschitz property of distance to a closed triangle set.
struct V {
 public double x,y,z;
 public V(double a,double b,double c){x=a;y=b;z=c;}
 public static V operator +(V a,V b){return new V(a.x+b.x,a.y+b.y,a.z+b.z);}
 public static V operator -(V a,V b){return new V(a.x-b.x,a.y-b.y,a.z-b.z);}
 public static V operator *(V a,double b){return new V(a.x*b,a.y*b,a.z*b);}
 public double Dot(V a){return x*a.x+y*a.y+z*a.z;}
 public V Cross(V b){return new V(y*b.z-z*b.y,z*b.x-x*b.z,x*b.y-y*b.x);}
 public double Norm(){return Math.Sqrt(Dot(this));}
 public double At(int a){return a==0?x:(a==1?y:z);}
 public override int GetHashCode(){unchecked{return (x.GetHashCode()*397^y.GetHashCode())*397^z.GetHashCode();}}
 public override bool Equals(object o){if(!(o is V))return false;V v=(V)o;return x==v.x&&y==v.y&&z==v.z;}
}
struct Tri {public V a,b,c;}
class Node {public V lo,hi; public int start,count,left=-1,right=-1;}
struct Hit {public double d;public V q;public int face;}
class GeometryCertificate {
 bool deferPlanar=false;
 double surfaceUncertainty=0;
 bool sourceIsNativeDifference=false;
 Tri[] tris;int[] order;List<Node> nodes=new List<Node>();
 static V Min(V a,V b){return new V(Math.Min(a.x,b.x),Math.Min(a.y,b.y),Math.Min(a.z,b.z));}
 static V Max(V a,V b){return new V(Math.Max(a.x,b.x),Math.Max(a.y,b.y),Math.Max(a.z,b.z));}
 static V Closest(V p,Tri t){
  V ab=t.b-t.a,ac=t.c-t.a,ap=p-t.a;
  if(ab.Cross(ac).Norm()<=1e-10){V u=Segment(p,t.a,t.b),v=Segment(p,t.a,t.c),w=Segment(p,t.b,t.c);return (p-u).Norm()<(p-v).Norm()?((p-u).Norm()<(p-w).Norm()?u:w):((p-v).Norm()<(p-w).Norm()?v:w);}
  double d1=ab.Dot(ap),d2=ac.Dot(ap);
  if(d1<=0&&d2<=0)return t.a;
  V bp=p-t.b;double d3=ab.Dot(bp),d4=ac.Dot(bp);if(d3>=0&&d4<=d3)return t.b;
  double vc=d1*d4-d3*d2;if(vc<=0&&d1>=0&&d3<=0)return t.a+ab*(d1/(d1-d3));
  V cp=p-t.c;double d5=ab.Dot(cp),d6=ac.Dot(cp);if(d6>=0&&d5<=d6)return t.c;
  double vb=d5*d2-d1*d6;if(vb<=0&&d2>=0&&d6<=0)return t.a+ac*(d2/(d2-d6));
  double va=d3*d6-d5*d4;if(va<=0&&(d4-d3)>=0&&(d5-d6)>=0)return t.b+(t.c-t.b)*((d4-d3)/((d4-d3)+(d5-d6)));
  double sum=va+vb+vc;
  if(Math.Abs(sum)<1e-30){V u=Segment(p,t.a,t.b),v=Segment(p,t.a,t.c),w=Segment(p,t.b,t.c);return (p-u).Norm()<(p-v).Norm()?((p-u).Norm()<(p-w).Norm()?u:w):((p-v).Norm()<(p-w).Norm()?v:w);}
  return t.a+ab*(vb/sum)+ac*(vc/sum);
 }
 static V Segment(V p,V a,V b){V d=b-a;double s=d.Dot(d);return s==0?a:a+d*Math.Max(0,Math.Min(1,(p-a).Dot(d)/s));}
 int Build(int start,int count){
  Node n=new Node{start=start,count=count,lo=new V(double.PositiveInfinity,double.PositiveInfinity,double.PositiveInfinity),hi=new V(double.NegativeInfinity,double.NegativeInfinity,double.NegativeInfinity)};
  for(int k=start;k<start+count;k++){Tri t=tris[order[k]];n.lo=Min(n.lo,Min(t.a,Min(t.b,t.c)));n.hi=Max(n.hi,Max(t.a,Max(t.b,t.c)));}
  int id=nodes.Count;nodes.Add(n);if(count<=12)return id;
  V e=n.hi-n.lo;int axis=e.x>=e.y&&e.x>=e.z?0:(e.y>=e.z?1:2);
  Array.Sort(order,start,count,Comparer<int>.Create((i,j)=>((tris[i].a.At(axis)+tris[i].b.At(axis)+tris[i].c.At(axis))).CompareTo(tris[j].a.At(axis)+tris[j].b.At(axis)+tris[j].c.At(axis))));
  int half=count/2;n.left=Build(start,half);n.right=Build(start+half,count-half);return id;
 }
 double Box(V p,Node n){double x=Math.Max(0,Math.Max(n.lo.x-p.x,p.x-n.hi.x)),y=Math.Max(0,Math.Max(n.lo.y-p.y,p.y-n.hi.y)),z=Math.Max(0,Math.Max(n.lo.z-p.z,p.z-n.hi.z));return x*x+y*y+z*z;}
 void Search(int id,V p,ref Hit h){Node n=nodes[id];if(Box(p,n)>h.d)return;
  if(n.left<0){for(int i=n.start;i<n.start+n.count;i++){int f=order[i];V q=Closest(p,tris[f]);double d=(p-q).Dot(p-q);if(d<h.d){h.d=d;h.q=q;h.face=f;}}return;}
  if(Box(p,nodes[n.left])<Box(p,nodes[n.right])){Search(n.left,p,ref h);Search(n.right,p,ref h);}else{Search(n.right,p,ref h);Search(n.left,p,ref h);}}
 [ThreadStatic] static Dictionary<V,Hit> memo;
 [ThreadStatic] static int lastFace;
 Hit Query(V p){
  if(memo==null)memo=new Dictionary<V,Hit>();Hit h;if(memo.TryGetValue(p,out h))return h;
  V seed=Closest(p,tris[lastFace]);h=new Hit{d=(p-seed).Dot(p-seed),q=seed,face=lastFace};
  Search(0,p,ref h);h.d=Math.Sqrt(h.d);lastFace=h.face;
  if(memo.Count>30000)memo.Clear();memo[p]=h;return h;
 }
 static V ReadV(BinaryReader r){return new V(r.ReadDouble(),r.ReadDouble(),r.ReadDouble());}
 static void WriteV(BinaryWriter w,V v){w.Write(v.x);w.Write(v.y);w.Write(v.z);}
 static Tri[] ReadMesh(string path){using(var r=new BinaryReader(File.OpenRead(path))){int n=r.ReadInt32();Tri[] t=new Tri[n];for(int i=0;i<n;i++)t[i]=new Tri{a=ReadV(r),b=ReadV(r),c=ReadV(r)};return t;}}
 // Exact triangle-to-triangle separation, not a maximum-distance witness.
 static double SegmentDistance(V p,V p1,V q,V q1){
  V u=p1-p,v=q1-q,w=p-q;double a=u.Dot(u),b=u.Dot(v),c=v.Dot(v),d=u.Dot(w),e=v.Dot(w),den=a*c-b*b;
  if(a<1e-25)return (p-Segment(p,q,q1)).Norm();if(c<1e-25)return (q-Segment(q,p,p1)).Norm();
  double s=den>1e-25?Math.Max(0,Math.Min(1,(b*e-c*d)/den)):0;
  double t=(b*s+e)/c;if(t<0){t=0;s=Math.Max(0,Math.Min(1,-d/a));}else if(t>1){t=1;s=Math.Max(0,Math.Min(1,(b-d)/a));}
  return (w+u*s-v*t).Norm();
 }
 static decimal DecimalValue(double value){return decimal.Parse(value.ToString("R",System.Globalization.CultureInfo.InvariantCulture),System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture);}
 static decimal TripleDecimal(V a,V b,V c){
  decimal ax=DecimalValue(a.x),ay=DecimalValue(a.y),az=DecimalValue(a.z),bx=DecimalValue(b.x),by=DecimalValue(b.y),bz=DecimalValue(b.z),cx=DecimalValue(c.x),cy=DecimalValue(c.y),cz=DecimalValue(c.z);
  return ax*(by*cz-bz*cy)+ay*(bz*cx-bx*cz)+az*(bx*cy-by*cx);
 }
 static bool RayTriRobust(V o,V v,Tri t,out double distance){
  V e1=t.b-t.a,e2=t.c-t.a,s=o-t.a;distance=0;decimal determinant=TripleDecimal(e1,v,e2);
  if(determinant==0)return false;
  decimal sign=determinant>0?1m:-1m,D=determinant*sign,U=TripleDecimal(s,v,e2)*sign,W=TripleDecimal(v,s,e1)*sign;
  if(U< -0.0000000001m*D||U>1.0000000001m*D||W< -0.0000000001m*D||U+W>1.0000000001m*D)return false;
  distance=(double)TripleDecimal(e2,s,e1)/(double)determinant;return distance>1e-8;
 }
 static bool RayTri(V o,V v,Tri t,out double distance){
  V e1=t.b-t.a,e2=t.c-t.a,h=v.Cross(e2);double a=e1.Dot(h);distance=0;
  if(Math.Abs(a)<=1e-12*e1.Norm()*h.Norm())return RayTriRobust(o,v,t,out distance);double f=1/a;V s=o-t.a;double u=f*s.Dot(h);if(u< -1e-10||u>1+1e-10)return false;
  V q=s.Cross(e1);double w=f*v.Dot(q);if(w< -1e-10||u+w>1+1e-10)return false;distance=f*e2.Dot(q);return distance>1e-8;
 }
 static bool EdgeCuts(V a,V b,Tri t){for(int k=0;k<3;k++)if(Math.Min(a.At(k),b.At(k))>Math.Max(t.a.At(k),Math.Max(t.b.At(k),t.c.At(k)))||Math.Max(a.At(k),b.At(k))<Math.Min(t.a.At(k),Math.Min(t.b.At(k),t.c.At(k))))return false;double d;return RayTri(a,b-a,t,out d)&&d<=1+1e-10;}
 static double TriDistance(Tri a,Tri b){
  V[] aa={a.a,a.b,a.c},bb={b.a,b.b,b.c};double d=double.PositiveInfinity;
  for(int i=0;i<3;i++){
   if(EdgeCuts(aa[i],aa[(i+1)%3],b)||EdgeCuts(bb[i],bb[(i+1)%3],a))return 0;
   d=Math.Min(d,(aa[i]-Closest(aa[i],b)).Norm());d=Math.Min(d,(bb[i]-Closest(bb[i],a)).Norm());
   for(int j=0;j<3;j++)d=Math.Min(d,SegmentDistance(aa[i],aa[(i+1)%3],bb[j],bb[(j+1)%3]));
  }return d;
 }
 double BoxDistance(V lo,V hi,Node n){double x=Math.Max(0,Math.Max(lo.x-n.hi.x,n.lo.x-hi.x)),y=Math.Max(0,Math.Max(lo.y-n.hi.y,n.lo.y-hi.y)),z=Math.Max(0,Math.Max(lo.z-n.hi.z,n.lo.z-hi.z));return Math.Sqrt(x*x+y*y+z*z);}
 void MinSearch(int node,Tri t,V lo,V hi,ref double distance){
  Node n=nodes[node];if(BoxDistance(lo,hi,n)>=distance)return;
  if(n.left<0){for(int k=n.start;k<n.start+n.count;k++){distance=Math.Min(distance,TriDistance(t,tris[order[k]]));if(distance<1e-8)return;}}
  else{MinSearch(n.left,t,lo,hi,ref distance);MinSearch(n.right,t,lo,hi,ref distance);}
 }
 double MinDistance(Tri t,double seed){V lo=Min(t.a,Min(t.b,t.c)),hi=Max(t.a,Max(t.b,t.c));MinSearch(0,t,lo,hi,ref seed);return seed;}
 static bool RayBox(V o,V v,Node n){double low=0,high=double.PositiveInfinity;for(int k=0;k<3;k++){double a=(n.lo.At(k)-o.At(k))/v.At(k),b=(n.hi.At(k)-o.At(k))/v.At(k);low=Math.Max(low,Math.Min(a,b));high=Math.Min(high,Math.Max(a,b));if(high<low)return false;}return true;}
 void RaySearch(int id,V o,V v,List<double> hits){Node n=nodes[id];if(!RayBox(o,v,n))return;if(n.left<0){for(int k=n.start;k<n.start+n.count;k++){double d;if(RayTri(o,v,tris[order[k]],out d))hits.Add(d);}}else{RaySearch(n.left,o,v,hits);RaySearch(n.right,o,v,hits);}}
 [ThreadStatic] static Dictionary<V,bool> insideMemo;
 bool Inside(V p){
  if(insideMemo==null)insideMemo=new Dictionary<V,bool>();bool cached;if(insideMemo.TryGetValue(p,out cached))return cached;
  V[] directions={new V(.81317,.34791,.46703),new V(-.51937,.72911,.44633),new V(.23753,-.59317,.76931)};int votes=0;
  foreach(V d in directions){var hits=new List<double>();RaySearch(0,p,d,hits);hits.Sort();int unique=0;double previous=double.NegativeInfinity;foreach(double h in hits)if(h-previous>1e-5){unique++;previous=h;}if(unique%2==1)votes++;}
  bool value=votes>=2;if(insideMemo.Count>30000)insideMemo.Clear();insideMemo[p]=value;return value;
 }
 struct Certificate {public int state;public Tri t;public double minimum,maximum;public V p,q;}
 class Work {public Tri t;public int depth;}
 static double Area(Tri t){return (t.b-t.a).Cross(t.c-t.a).Norm()*.5;}
 List<Certificate> Certify(Tri original,double threshold){
  double lowerThreshold=threshold+surfaceUncertainty,upperThreshold=threshold-surfaceUncertainty;
  var result=new List<Certificate>();var pending=new Stack<Work>();pending.Push(new Work{t=original,depth=0});int count=0;
  while(pending.Count>0){var work=pending.Pop();Tri t=work.t;double area=Area(t);
   V center=(t.a+t.b+t.c)*(1.0/3);V[] pp={t.a,t.b,t.c,center};Hit[] hh=new Hit[4];double maximum=0,upper=double.PositiveInfinity;int witness=0;
   for(int i=0;i<4;i++){hh[i]=Query(pp[i]);if(hh[i].d>maximum){maximum=hh[i].d;witness=i;}}
   double minimum=MinDistance(t,hh[3].d);bool inside=!sourceIsNativeDifference&&hh[3].d>.001&&Inside(center);int state=-1;
   if(minimum>.001&&inside)state=3;
   else if(minimum>=lowerThreshold&&maximum>lowerThreshold&&!inside)state=1;
   else {
    for(int i=0;i<4;i++){Tri target=tris[hh[i].face];upper=Math.Min(upper,Math.Max((t.a-Closest(t.a,target)).Norm(),Math.Max((t.b-Closest(t.b,target)).Norm(),(t.c-Closest(t.c,target)).Norm())));}
    if(upper<=upperThreshold)state=0;
   }
   if(state<0&&minimum>=upperThreshold&&maximum<=lowerThreshold)state=2;
   if(state<0&&deferPlanar&&maximum<=lowerThreshold)state=2;
   if(state<0&&(work.depth>=24||++count>512))state=2;
   if(state>=0){result.Add(new Certificate{state=state,t=t,minimum=minimum,maximum=maximum,p=pp[witness],q=hh[witness].q});continue;}
   // Split by the sampled threshold first. Every resulting triangle is still
   // checked against an exact minimum/convex upper bound in the next pass.
   double splitThreshold=threshold;
   double dmin=Math.Min(hh[0].d,Math.Min(hh[1].d,hh[2].d)),dmax=Math.Max(hh[0].d,Math.Max(hh[1].d,hh[2].d));
   if(dmin<upperThreshold&&dmax>upperThreshold)splitThreshold=upperThreshold;
   else if(dmin<lowerThreshold&&dmax>lowerThreshold)splitThreshold=lowerThreshold;
   double[] values=new double[3];bool pos=false,neg=false;
   for(int i=0;i<3;i++){values[i]=(!sourceIsNativeDifference&&hh[i].d>.001&&Inside(pp[i])?-hh[i].d:hh[i].d)-splitThreshold;pos|=values[i]>1e-5;neg|=values[i]<-1e-5;}
   if(pos&&neg){
    for(int sign=-1;sign<=1;sign+=2){var polygon=new List<V>();for(int i=0;i<3;i++){int j=(i+1)%3;double a=sign*values[i],b=sign*values[j];if(a>=0)polygon.Add(pp[i]);if((a>=0)!=(b>=0))polygon.Add(pp[i]+(pp[j]-pp[i])*(a/(a-b)));}
     for(int i=1;i+1<polygon.Count;i++)pending.Push(new Work{t=new Tri{a=polygon[0],b=polygon[i],c=polygon[i+1]},depth=work.depth+1});
    }
   }else{
    double ab=(t.a-t.b).Norm(),bc=(t.b-t.c).Norm(),ca=(t.c-t.a).Norm();Tri a,b;
    if(ab>=bc&&ab>=ca){V m=(t.a+t.b)*.5;a=new Tri{a=t.a,b=m,c=t.c};b=new Tri{a=m,b=t.b,c=t.c};}
    else if(bc>=ca){V m=(t.b+t.c)*.5;a=new Tri{a=t.a,b=t.b,c=m};b=new Tri{a=t.a,b=m,c=t.c};}
    else{V m=(t.c+t.a)*.5;a=new Tri{a=t.a,b=t.b,c=m};b=new Tri{a=m,b=t.b,c=t.c};}
    pending.Push(new Work{t=a,depth=work.depth+1});pending.Push(new Work{t=b,depth=work.depth+1});
   }
  }return result;
 }
 static int Main(string[]args){try{
  var engine=new GeometryCertificate();engine.tris=ReadMesh(args[0]);engine.order=new int[engine.tris.Length];for(int i=0;i<engine.order.Length;i++)engine.order[i]=i;engine.Build(0,engine.order.Length);
  if(args[3]=="bounds"){
   Tri[] queryTriangles=ReadMesh(args[1]);int count=queryTriangles.Length;
   double[] minima=new double[count],uppers=new double[count],maxima=new double[count];
   V[] ppout=new V[count],qqout=new V[count];int[] faceout=new int[count];
   Parallel.For(0,count,new ParallelOptions{MaxDegreeOfParallelism=2},i=>{
    Tri t=queryTriangles[i];V[] pp={t.a,t.b,t.c,(t.a+t.b+t.c)*(1.0/3)};
    Hit[] hh=new Hit[4];double maximum=-1,upper=double.PositiveInfinity;int witness=0;
    for(int j=0;j<4;j++){hh[j]=engine.Query(pp[j]);if(hh[j].d>maximum){maximum=hh[j].d;witness=j;}}
    for(int j=0;j<4;j++){Tri target=engine.tris[hh[j].face];upper=Math.Min(upper,Math.Max((t.a-Closest(t.a,target)).Norm(),Math.Max((t.b-Closest(t.b,target)).Norm(),(t.c-Closest(t.c,target)).Norm())));}
    minima[i]=engine.MinDistance(t,hh[3].d);uppers[i]=upper;maxima[i]=maximum;ppout[i]=pp[witness];qqout[i]=hh[witness].q;faceout[i]=hh[witness].face;
   });
   using(var wr=new BinaryWriter(File.Create(args[2]))){wr.Write(count);for(int i=0;i<count;i++){wr.Write(minima[i]);wr.Write(uppers[i]);wr.Write(maxima[i]);wr.Write(faceout[i]);WriteV(wr,ppout[i]);WriteV(wr,qqout[i]);}}return 0;
  }
  if(args[3]=="points"){
   V[] points;using(var rd=new BinaryReader(File.OpenRead(args[1]))){int n=rd.ReadInt32();points=new V[n];for(int i=0;i<n;i++)points[i]=ReadV(rd);}
   Hit[] answers=new Hit[points.Length];Parallel.For(0,points.Length,new ParallelOptions{MaxDegreeOfParallelism=2},i=>answers[i]=engine.Query(points[i]));
   using(var wr=new BinaryWriter(File.Create(args[2]))){wr.Write(points.Length);foreach(Hit h in answers){wr.Write(h.d);wr.Write(h.face);WriteV(wr,h.q);}}return 0;
  }
  engine.deferPlanar=args.Length>4&&args[4]=="fast";
  engine.sourceIsNativeDifference=args.Length>5&&args[5]=="native-csg-difference";
  engine.surfaceUncertainty=args.Length>6?double.Parse(args[6],System.Globalization.CultureInfo.InvariantCulture):0;
  if(engine.surfaceUncertainty<0)throw new ArgumentException("Negative uncertainty");
  Tri[] source=ReadMesh(args[1]);double tolerance=double.Parse(args[3],System.Globalization.CultureInfo.InvariantCulture);
  using(var w=new BinaryWriter(File.Create(args[2]))){w.Write(0);int total=0;const int chunk=100;
   for(int start=0;start<source.Length;start+=chunk){int n=Math.Min(chunk,source.Length-start);var rr=new List<Certificate>[n];int offset=start;Parallel.For(0,n,new ParallelOptions{MaxDegreeOfParallelism=2},i=>rr[i]=engine.Certify(source[offset+i],tolerance));int[] counts=new int[4];
    for(int i=0;i<n;i++)foreach(var r in rr[i]){w.Write(offset+i);w.Write(r.state);WriteV(w,r.t.a);WriteV(w,r.t.b);WriteV(w,r.t.c);w.Write(r.minimum);w.Write(r.maximum);WriteV(w,r.p);WriteV(w,r.q);counts[r.state]++;total++;}
    w.Flush();Console.WriteLine("CERTIFY "+(start+n)+"/"+source.Length+" within="+counts[0]+" over="+counts[1]+" unresolved="+counts[2]+" inside="+counts[3]);
   }w.BaseStream.Position=0;w.Write(total);
  }return 0;
 }catch(Exception e){Console.Error.WriteLine(e);return 1;}}
}

