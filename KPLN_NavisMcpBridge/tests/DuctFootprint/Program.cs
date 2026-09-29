using KPLN_NavisMcpBridge;

var tests = new (string, Action)[] {
    ("empty auxiliary fragment is skipped", () => {
        var chosen=FootprintMeshValidation.Select(new List<Point3Dto>(),new List<Point3Dto>(),Identity(),
            Bounds(P(0,0,0),P(0,0,1)),out var reason);
        Assert(chosen != null && chosen.Count == 0 && reason == "no-triangle-mesh-skipped",reason);
    }),
    ("conservative world bounds enclose rotated irregular mesh", () => {
        var row = Shift(Rotate(Box(188,473,189),.63,0),7212103,1041196,48);
        var matrix = Identity(); matrix[12]=7212103;matrix[13]=1041196;matrix[14]=48;
        var chosen=FootprintMeshValidation.Select(row,Box(188,473,189),matrix,
            Bounds(P(7212100,1041193,47),P(7212106,1041199,49)),out var reason);
        Assert(chosen==row,reason);
        Check(Measure(chosen.ToList(),Shift(Box(3000,3000,50),7212103,1041196,48),P(7212103,1041196,48)),188,473,188*473);
    }),
    ("wrong instance outside world box rejected", () => {
        var row=Shift(Box(188,473,189),10,0,0);var m=Identity();m[12]=10;
        Assert(FootprintMeshValidation.Select(row,row,m,Bounds(P(-1,-1,-1),P(1,1,1)),out _)==null,"foreign mesh accepted");
    }),
    ("inward mismatch must not hide one outside vertex", () => {
        var row=Box(188,473,189);row[0]=P(5,0,0);
        Assert(FootprintMeshValidation.Select(row,row,Identity(),Bounds(P(-1,-1,-1),P(1,1,1)),out _)==null,"outside vertex accepted");
    }),
    ("ambiguous affine rotation rejected", () => {
        var a=Box(188,473,189);var b=Rotate(a,.5,0);
        Assert(FootprintMeshValidation.Select(a,b,Identity(),Bounds(P(-2,-2,-2),P(2,2,2)),out var why)==null && why=="ambiguous-world-transform",why);
    }),
    ("projective transform rejected even with contained points", () => {
        var a=Box(188,473,189);var m=Identity();m[3]=2;m[12]=2;
        Assert(FootprintMeshValidation.Select(a,a,m,Bounds(P(-2,-2,-2),P(2,2,2)),out _)==null,"projective accepted");
    }),
    ("nonfinite transform rejected", () => {
        var a=Box(188,473,189);var m=Identity();m[0]=double.NaN;
        Assert(FootprintMeshValidation.Select(a,a,m,Bounds(P(-2,-2,-2),P(2,2,2)),out _)==null,"nonfinite accepted");
    }),
    ("column affine convention supported", () => {
        var a=Shift(Box(188,473,189),10,0,0);var m=Identity();m[3]=10;
        Assert(FootprintMeshValidation.Select(Box(188,473,189),a,m,Bounds(P(9,-1,-1),P(11,1,1)),out var why)==a,why);
    }),
    ("short vertical 188x473 length189", () => Check(Measure(Box(188,473,189), Box(3000,3000,50)), 188,473,188*473)),
    ("round vertical diameter100", () => CheckRound(Measure(Cylinder(100,200,48), Box(3000,3000,50)),100)),
    ("rotated in plan", () => Check(Measure(Rotate(Box(188,473,189),.63,0), Box(3000,3000,50)),188,473,188*473)),
    ("sloped ceiling", () => Check(Measure(Rotate(Box(188,473,189),.63,.42), Rotate(Box(3000,3000,50),.63,.42)),188,473,188*473)),
    ("large coordinates", () => Check(Measure(Shift(Rotate(Box(188,473,189),.63,0),7212103,1041196,48),
        Shift(Box(3000,3000,50),7212103,1041196,48), P(7212103,1041196,48)),188,473,188*473)),
    ("longitudinal larger footprint", () => Check(Measure(Box(188,2000,473), Box(3000,3000,50)),188,2000,188*2000)),
    ("partial overlap", () => Check(Measure(Box(188,473,189),Shift(Box(3000,3000,50),1500/304.8,0,0)),94,473,94*473)),
    ("ceiling hole preserves missing area", () => {
        var ceiling = Shift(Box(1450,3000,50),-775/304.8,0,0);
        ceiling.AddRange(Shift(Box(1450,3000,50),775/304.8,0,0));
        ceiling.AddRange(Shift(Box(100,1450,50),0,-775/304.8,0));
        ceiling.AddRange(Shift(Box(100,1450,50),0,775/304.8,0));
        Check(Measure(Box(188,473,189),ceiling),188,473,188*473-100*100);
    }),
    ("no overlap", () => Assert(!Measure(Box(188,473,189),Shift(Box(3000,3000,50),100,0,0)).Usable,"false intersection")),
    ("parallel displaced", () => Assert(!Measure(Box(188,473,189),Shift(Box(3000,3000,50),0,0,10)).Usable,"false slice")),
    ("missing mesh", () => Assert(!Measure(new(),Box(3000,3000,50)).Usable,"empty accepted")),
    ("ambiguous ceiling", () => Assert(!Measure(Box(188,473,189),Box(500,500,500)).Usable,"cube plane guessed")),
    ("duplicated export faces", () => {
        var ceiling=Box(3000,3000,50); ceiling.AddRange(Box(3000,3000,50));
        Check(Measure(Box(188,473,189),ceiling),188,473,188*473);
    }),
    ("open duct ends", () => Check(Measure(Box(188,473,189).Skip(12).ToList(),Box(3000,3000,50)),188,473,188*473)),
    ("horizontal ceiling plane", () => CheckPlane(DuctCeilingFootprint.MeasurePlane(Box(3000,3000,50)), true)),
    ("sloped ceiling plane", () => CheckPlane(DuctCeilingFootprint.MeasurePlane(Rotate(Box(3000,3000,50),.63,.42)), false)),
    ("ambiguous standalone surface", () => Assert(!DuctCeilingFootprint.MeasurePlane(Box(500,500,500)).Usable,"cube plane guessed")),
};
foreach(var (name,test) in tests) { test(); Console.WriteLine("PASS "+name); }
Console.WriteLine($"{tests.Length} duct footprint tests passed.");
static ClashFootprintDto Measure(List<Point3Dto> duct,List<Point3Dto> ceiling,Point3Dto p=null) => DuctCeilingFootprint.Measure(duct,ceiling,p??P(0,0,0));
static void Assert(bool ok,string message) { if(!ok) throw new Exception(message); }
static void Check(ClashFootprintDto r,double a,double b,double area) {
    Assert(r.Usable,r.Reason);
    Assert(Math.Abs(r.MinEdge*304.8-a)<.01,"min edge "+r.MinEdge*304.8);
    Assert(Math.Abs(r.MaxEdge*304.8-b)<.01,"max edge "+r.MaxEdge*304.8);
    Assert(Math.Abs(r.Area*304.8*304.8-area)<1,"area "+r.Area*304.8*304.8);
}
static void CheckRound(ClashFootprintDto r,double diameter) {
    Assert(r.Usable,r.Reason);
    Assert(Math.Abs(r.MinEdge*304.8-diameter)<=diameter*.02,"round min edge "+r.MinEdge*304.8);
    Assert(Math.Abs(r.MaxEdge*304.8-diameter)<=diameter*.02,"round max edge "+r.MaxEdge*304.8);
    var expected=Math.PI*diameter*diameter/4;
    Assert(Math.Abs(r.Area*304.8*304.8-expected)<=expected*.02,"round area "+r.Area*304.8*304.8);
}
static void CheckPlane(SurfacePlaneDto r,bool horizontal) {
    Assert(r.Usable,r.Reason);
    var length=Math.Sqrt(r.Normal.X*r.Normal.X+r.Normal.Y*r.Normal.Y+r.Normal.Z*r.Normal.Z);
    Assert(Math.Abs(length-1)<1e-6,"normal is not unit");
    Assert(r.DominantAreaRatio>=2,"ambiguous dominant plane");
    if(horizontal) Assert(Math.Abs(Math.Abs(r.Normal.Z)-1)<1e-6,"horizontal normal expected");
    else Assert(Math.Abs(r.Normal.Z)<.999,"sloped normal expected");
}
static Point3Dto P(double x,double y,double z)=>new Point3Dto{X=x,Y=y,Z=z};
static BoundDto Bounds(Point3Dto min,Point3Dto max)=>new BoundDto{Min=min,Max=max};
static double[] Identity()=>new double[]{1,0,0,0,0,1,0,0,0,0,1,0,0,0,0,1};
static List<Point3Dto> Shift(List<Point3Dto> p,double x,double y,double z)=>p.Select(v=>P(v.X+x,v.Y+y,v.Z+z)).ToList();
static List<Point3Dto> Rotate(List<Point3Dto> p,double az,double tilt)=>p.Select(v=> {
    var x=v.X*Math.Cos(az)-v.Y*Math.Sin(az);var y=v.X*Math.Sin(az)+v.Y*Math.Cos(az);
    return P(x,y*Math.Cos(tilt)-v.Z*Math.Sin(tilt),y*Math.Sin(tilt)+v.Z*Math.Cos(tilt));
}).ToList();
static List<Point3Dto> Box(double x,double y,double z) {
    var p=new List<Point3Dto>();
    for(int i=0;i<8;i++)p.Add(P(((i&1)==0?-.5:.5)*x/304.8,((i&2)==0?-.5:.5)*y/304.8,((i&4)==0?-.5:.5)*z/304.8));
    int[] ids={0,2,3,0,3,1,4,5,7,4,7,6,0,1,5,0,5,4,2,6,7,2,7,3,0,4,6,0,6,2,1,3,7,1,7,5};
    return ids.Select(i=>p[i]).ToList();
}
static List<Point3Dto> Cylinder(double diameter,double length,int segments) {
    var points=new List<Point3Dto>();var radius=diameter/2/304.8;var half=length/2/304.8;
    for(int i=0;i<segments;i++) {
        var a=2*Math.PI*i/segments;var b=2*Math.PI*(i+1)/segments;
        var p0=P(radius*Math.Cos(a),radius*Math.Sin(a),-half);var p1=P(radius*Math.Cos(b),radius*Math.Sin(b),-half);
        var p2=P(radius*Math.Cos(a),radius*Math.Sin(a),half);var p3=P(radius*Math.Cos(b),radius*Math.Sin(b),half);
        points.AddRange(new[]{p0,p1,p3,p0,p3,p2,P(0,0,-half),p1,p0,P(0,0,half),p2,p3});
    }
    return points;
}
namespace KPLN_NavisMcpBridge {
    public class Point3Dto { public double X,Y,Z; }
    public class BoundDto { public Point3Dto Min,Max; }
}
