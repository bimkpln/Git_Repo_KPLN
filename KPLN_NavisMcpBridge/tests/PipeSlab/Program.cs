using KPLN_NavisMcpBridge;

var tests = new (string, Action)[] {
    ("long vertical pipe", () => Normal(Cylinder(100,1000),Box(3000,3000,200))),
    ("short pipe crosses one slab face", () => Normal(Shift(Cylinder(100,60),0,0,-100),Box(3000,3000,200))),
    ("very short pipe axis is not largest dimension", () => Normal(Shift(Cylinder(100,10),0,0,-100),Box(3000,3000,200))),
    ("uncapped straight pipe", () => Normal(Cylinder(100,500,false),Box(3000,3000,200))),
    ("intermediate axial mesh rings", () => Normal(SegmentedCylinder(100,500,4),Box(3000,3000,200))),
    ("joined fragments with different tessellation", () => {
        var p=Shift(Cylinder(100,250,false,16),0,0,-125);
        p.AddRange(Shift(Cylinder(100,250,false,32),0,0,125));
        Normal(p,Box(3000,3000,200));
    }),
    ("densely tessellated cylinder", () => Normal(Cylinder(100,500,true,512),Box(3000,3000,200))),
    ("slightly noisy cylinder mesh", () => Normal(Noisy(SegmentedCylinder(100,500,4)),Box(3000,3000,200))),
    ("per-triangle vertex noise is welded", () => Normal(OccurrenceNoise(SegmentedCylinder(100,500,4)),Box(3000,3000,200))),
    ("noisy uncapped cylinder axis fit", () => Normal(Noisy(SegmentedCylinder(100,500,4,false)),Box(3000,3000,200))),
    ("separated coaxial pieces rejected", () => {
        var p=Shift(Cylinder(100,200),0,0,-125);p.AddRange(Shift(Cylinder(100,200),0,0,125));
        Check(!Measure(p,Box(3000,3000,200)).Usable,"gapped pipe pieces accepted");
    }),
    ("hollow pipe", () => {
        var p=Cylinder(108,500,false);p.AddRange(Cylinder(100,500,false));Normal(p,Box(3000,3000,200));
    }),
    ("pipe within slab not mistaken for face passage", () => {
        var r=Measure(Cylinder(100,60),Box(3000,3000,200));Check(r.Usable && !r.FullSectionCovered && r.Sections.Count==0,r.Reason);
    }),
    ("horizontal pipe within slab", () => {
        var r=Measure(Rotate(Cylinder(100,500),Math.PI/2),Box(3000,3000,200));
        Check(r.Usable && Math.Abs(Dot(r.PipeAxis,r.SlabNormal))<1e-6 && !r.FullSectionCovered,r.Reason);
    }),
    ("sloped slab and normal pipe", () => Normal(Rotate(Cylinder(100,1000),.47),Rotate(Box(3000,3000,200),.47))),
    ("oblique pipe has actual angle", () => {
        var r=Measure(Rotate(Cylinder(100,1000),.4),Box(3000,3000,200));
        Check(r.Usable && r.FullSectionCovered && Math.Abs(Math.Abs(Dot(r.PipeAxis,r.SlabNormal))-Math.Cos(.4))<1e-6,r.Reason);
    }),
    ("outer slab edge", () => {
        var r=Measure(Cylinder(100,1000),Shift(Box(3000,3000,200),1500,0,0));
        Check(r.Usable && r.EdgeContact && !r.FullSectionCovered,r.Reason);
    }),
    ("hole edge", () => {
        var slab=RingSlab();
        var r=Measure(Cylinder(150,1000),slab);Check(r.Usable && r.EdgeContact && !r.FullSectionCovered,r.Reason);
    }),
    ("far slab edge not contact", () => Normal(Shift(Cylinder(100,500),1400,0,0),Box(3000,3000,200))),
    ("edge above pipe endpoints not contact", () => {
        var r=Measure(Shift(Cylinder(100,30),1500,0,1000),Box(3000,3000,200));Check(r.Usable && !r.EdgeContact && !r.FullSectionCovered,r.Reason);
    }),
    ("incomplete slab rejected", () => Check(!Measure(Cylinder(100,500),Box(3000,3000,200).Skip(3).ToList()).Usable,"open slab accepted")),
    ("duplicate slab triangles tolerated", () => {var s=Box(3000,3000,200);s.AddRange(Box(3000,3000,200));Normal(Cylinder(100,500),s);}),
    ("box cannot masquerade as pipe", () => Check(!Measure(Box(100,100,500),Box(3000,3000,200)).Usable,"box accepted")),
    ("ellipse cannot masquerade as circle", () => Check(!Measure(Cylinder(100,500).Select(p=>P(p.X*1.4,p.Y,p.Z)).ToList(),Box(3000,3000,200)).Usable,"ellipse accepted")),
    ("incomplete pipe wall rejected", () => Check(!Measure(Cylinder(100,500,false).Skip(6).ToList(),Box(3000,3000,200)).Usable,"incomplete pipe accepted")),
    ("multi-axis element rejected", () => {var p=Cylinder(100,500);p.AddRange(Rotate(Cylinder(100,500),.5));Check(!Measure(p,Box(3000,3000,200)).Usable,"multi-axis accepted");}),
    ("world coordinates", () => {
        var p=Shift(Rotate(Cylinder(100,500),.4),7212103*304.8,1041196*304.8,48*304.8);
        var s=Shift(Rotate(Box(3000,3000,200),.4),7212103*304.8,1041196*304.8,48*304.8);
        var r=DuctCeilingFootprint.MeasurePipeSlab(p,s,P(7212103,1041196,48));Check(r.Usable && r.FullSectionCovered && !r.EdgeContact,r.Reason);
    }),
    ("ambiguous slab cube", () => Check(!Measure(Cylinder(100,1000),Box(500,500,500)).Usable,"cube guessed")),
};
foreach(var (name,test) in tests) { test(); Console.WriteLine("PASS "+name); }
Console.WriteLine($"{tests.Length} pipe slab tests passed.");
static double Dot(Point3Dto a,Point3Dto b)=>a.X*b.X+a.Y*b.Y+a.Z*b.Z;
static void Normal(List<Point3Dto> p,List<Point3Dto> s) {var r=Measure(p,s);Check(r.Usable && r.FullSectionCovered && !r.EdgeContact,r.Reason);Check(Math.Abs(Dot(r.PipeAxis,r.SlabNormal))>1-1e-6,"wrong axis");}
static PipeSlabDto Measure(List<Point3Dto> p,List<Point3Dto> s)=>DuctCeilingFootprint.MeasurePipeSlab(p,s,P(0,0,0));
static void Check(bool ok,string message){if(!ok)throw new Exception(message??"failed assertion");}
static Point3Dto P(double x,double y,double z)=>new Point3Dto{X=x,Y=y,Z=z};
static List<Point3Dto> Shift(List<Point3Dto> p,double x,double y,double z)=>p.Select(v=>P(v.X+x/304.8,v.Y+y/304.8,v.Z+z/304.8)).ToList();
static List<Point3Dto> Rotate(List<Point3Dto> p,double t)=>p.Select(v=>P(v.X,v.Y*Math.Cos(t)-v.Z*Math.Sin(t),v.Y*Math.Sin(t)+v.Z*Math.Cos(t))).ToList();
static List<Point3Dto> Box(double x,double y,double z) {
    var p=new List<Point3Dto>();for(int i=0;i<8;i++)p.Add(P(((i&1)==0?-.5:.5)*x/304.8,((i&2)==0?-.5:.5)*y/304.8,((i&4)==0?-.5:.5)*z/304.8));
    return new[]{0,2,3,0,3,1,4,5,7,4,7,6,0,1,5,0,5,4,2,6,7,2,7,3,0,4,6,0,6,2,1,3,7,1,7,5}.Select(i=>p[i]).ToList();
}
static List<Point3Dto> Cylinder(double diameter,double length,bool capped=true,int count=32) {
    var p=new List<Point3Dto>();var r=diameter/2/304.8;var h=length/2/304.8;
    for(int i=0;i<count;i++) {var a=2*Math.PI*i/count;var b=2*Math.PI*(i+1)/count;
        var p0=P(r*Math.Cos(a),r*Math.Sin(a),-h);var p1=P(r*Math.Cos(b),r*Math.Sin(b),-h);var p2=P(r*Math.Cos(a),r*Math.Sin(a),h);var p3=P(r*Math.Cos(b),r*Math.Sin(b),h);
        p.AddRange(new[]{p0,p1,p3,p0,p3,p2});if(capped)p.AddRange(new[]{P(0,0,-h),p1,p0,P(0,0,h),p2,p3});
    }return p;
}
static List<Point3Dto> SegmentedCylinder(double diameter,double length,int segments,bool capped=true) {
    var p=new List<Point3Dto>();var r=diameter/2/304.8;var h=length/2/304.8;int count=32;
    for(int s=0;s<segments;s++) {var z0=-h+2*h*s/segments;var z1=-h+2*h*(s+1)/segments;
        for(int i=0;i<count;i++) {var a=2*Math.PI*i/count;var b=2*Math.PI*(i+1)/count;
            var p0=P(r*Math.Cos(a),r*Math.Sin(a),z0);var p1=P(r*Math.Cos(b),r*Math.Sin(b),z0);
            var p2=P(r*Math.Cos(a),r*Math.Sin(a),z1);var p3=P(r*Math.Cos(b),r*Math.Sin(b),z1);
            p.AddRange(new[]{p0,p1,p3,p0,p3,p2});
            if(capped&&s==0)p.AddRange(new[]{P(0,0,-h),p1,p0});
            if(capped&&s==segments-1)p.AddRange(new[]{P(0,0,h),p2,p3});
        }
    }return p;
}
static List<Point3Dto> Noisy(List<Point3Dto> points) => points.Select(p => {
    var radius=Math.Sqrt(p.X*p.X+p.Y*p.Y);if(radius<1e-9)return p;
    var angle=Math.Atan2(p.Y,p.X);var delta=.02/304.8*Math.Sin(angle*7+p.Z*13);
    return P(p.X*(radius+delta)/radius,p.Y*(radius+delta)/radius,p.Z);
}).ToList();
static List<Point3Dto> OccurrenceNoise(List<Point3Dto> points) => points.Select((p,i) => {
    var d=.001/304.8;
    return P(p.X+d*Math.Sin(i*1.73),p.Y+d*Math.Sin(i*2.31),p.Z+d*Math.Sin(i*3.17));
}).ToList();
static List<Point3Dto> RingSlab() {
    var vertices=new List<Point3Dto>();
    foreach(var z in new[]{-100.0,100.0}) foreach(var r in new[]{1500.0,50.0})
        foreach(var xy in new[]{(-1,-1),(1,-1),(1,1),(-1,1)}) vertices.Add(P(xy.Item1*r/304.8,xy.Item2*r/304.8,z/304.8));
    var mesh=new List<Point3Dto>();
    void Quad(int a,int b,int c,int d){mesh.AddRange(new[]{a,b,c,a,c,d}.Select(i=>vertices[i]));}
    for(int i=0;i<4;i++){int j=(i+1)%4;Quad(i,j,4+j,4+i);Quad(8+i,12+i,12+j,8+j);Quad(i,8+i,8+j,j);Quad(4+i,4+j,12+j,12+i);}
    return mesh;
}
namespace KPLN_NavisMcpBridge { public class Point3Dto { public double X,Y,Z; } public class BoundDto {public Point3Dto Min,Max;} }
