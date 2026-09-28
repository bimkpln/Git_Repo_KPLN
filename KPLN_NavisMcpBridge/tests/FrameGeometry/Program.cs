using KPLN_NavisMcpBridge;

var tests = new (string, Action)[]
{
    ("COM array lower bound", () => {
        var a = Array.CreateInstance(typeof(int), new[] { 3 }, new[] { 1 });
        a.SetValue(2, 1); a.SetValue(30, 2); a.SetValue(4, 3);
        Assert(GeometryInstancePath.Read(a) == GeometryInstancePath.Read(new[] { 2, 30, 4 }), "lower bound changed path");
    }),
    ("path key no concatenation collision", () => {
        Assert(GeometryInstancePath.Read(new[] { 1, 23 }) != GeometryInstancePath.Read(new[] { 12, 3 }), "path collision");
    }),
    ("exact path not shared tail or prefix", () => {
        var paths = new[] { new[] { 1, 2, 3 }, new[] { 9, 2, 3 }, new[] { 1, 2, 3, 4 }, new[] { 1, 2 }, new[] { 1, 2, 3 } };
        var selected = GeometryInstancePath.Select(paths, "1/2/3", p => p, out var skipped);
        Assert(selected.Count == 2 && skipped == 3, "foreign instance selected or own fragment lost");
    }),
    ("filter before triangle budget", () => {
        var paths = Enumerable.Range(1, 1000).Select(i => new[] { 1, i, 99 }).ToArray();
        int callbacks = 0;
        var selected = GeometryInstancePath.Select(paths, "1/500/99", p => p, out var skipped);
        var mesh = new List<Point3Dto>();
        foreach (var p in selected) { callbacks++; mesh.AddRange(Rotate(Box(8, .2, .5), -.8)); }
        Assert(skipped == 999 && callbacks == 1 && mesh.Count < 30000, "foreign instances consumed budget");
        CheckAxis(mesh, new[] { Math.Cos(-.8), Math.Sin(-.8), 0 });
    }),
    ("unknown instance rejected", () => ExpectFailure(() => GeometryInstancePath.Select(new[] { new[] { 1, 2 } }, "2/1", p => p, out _))),
    ("invalid instance arrays rejected", () => {
        foreach (var data in new object[] { null, Array.Empty<int>(), new int[2, 2], new object[] { 1, null }, new[] { 1.5 }, new[] { "1" } })
            ExpectFailure(() => GeometryInstancePath.Read(data));
    }),
    ("unreadable fragment cannot be silently dropped", () => ExpectFailure(() =>
        GeometryInstancePath.Select(new object[] { new[] { 1, 2 }, null }, "1/2", p => p, out _))),
    ("fragment box union", () => {
        var a = Box(8, .2, .5); var b = a.Select(p => P(p.X + 30, p.Y, p.Z)).ToList();
        var box = MeshAxisFitter.UnionBounds(new[] { Bounds(a), Bounds(b) });
        Assert(box.Min.X == -4 && box.Max.X == 34, "wrong fragment union");
        Assert(MeshAxisFitter.UnionBounds(new BoundDto[] { Bounds(a), null }) == null, "partial box accepted");
    }),
    ("instance box not shared node box", () => {
        var mesh = Rotate(Box(8, .2, .5), -.8);
        var selectedBox = Bounds(mesh);
        var sharedBox = MeshAxisFitter.UnionBounds(new[] { selectedBox, Bounds(Box(100, 100, 10)) });
        Assert(!MeshAxisFitter.FitWorld(mesh, mesh, sharedBox).Usable, "shared bounds unexpectedly matched");
        Assert(MeshAxisFitter.FitWorld(mesh, mesh, selectedBox).Usable, "selected instance rejected");
    }),
    ("world X", () => CheckAxis(Box(8, .2, .5), new[] { 1d, 0, 0 })),
    ("diagonal negative XY", () => CheckAxis(Rotate(Box(8, .2, .5), -.8), new[] { Math.Cos(-.8), Math.Sin(-.8), 0 })),
    ("same AABB different axes", () => {
        var a = Fit(Rotate(Box(8, .2, .5), Math.PI / 4));
        var b = Fit(Rotate(Box(8, .2, .5), -Math.PI / 4));
        Assert(Math.Abs(a.Axis.X * b.Axis.X + a.Axis.Y * b.Axis.Y) < 1e-8, "axes collapsed to AABB");
    }),
    ("ambiguous transform rejected", () => {
        var a = Rotate(Box(8, .2, .5), Math.PI / 4);
        var b = Rotate(Box(8, .2, .5), -Math.PI / 4);
        var result = MeshAxisFitter.FitWorld(a, b, Bounds(a));
        Assert(!result.Usable && result.Reason == "ambiguous-world-transform", "ambiguous transform accepted");
    }),
    ("identity transform equivalent", () => {
        var a = Box(8, .2, .5);
        Assert(MeshAxisFitter.FitWorld(a, a, Bounds(a)).Usable, "identity rejected");
    }),
    ("select world transform", () => {
        var a = Rotate(Box(8, .2, .5), -.8);
        var b = a.Select(p => P(p.X + 7212103, p.Y + 1041196, p.Z + 48)).ToList();
        Assert(MeshAxisFitter.FitWorld(a, b, Bounds(b)).Source == "mesh-surface-pca-column", "wrong convention");
        Assert(MeshAxisFitter.FitWorld(b, a, Bounds(b)).Source == "mesh-surface-pca-row", "wrong convention");
    }),
    ("bounds mismatch rejected", () => {
        var a = Box(8, .2, .5);
        Assert(!MeshAxisFitter.FitWorld(a, a, Bounds(Box(12, 4, 4))).Usable, "mismatch accepted");
    }),
    ("inclined XYZ", () => CheckAxis(Rotate(Box(8, .2, .5), .7, -.4),
        new[] { Math.Cos(.7) * Math.Cos(-.4), Math.Sin(.7) * Math.Cos(-.4), Math.Sin(-.4) })),
    ("vertical", () => CheckAxis(Box(.2, .5, 8), new[] { 0d, 0, 1 })),
    ("large world translation", () => CheckAxis(Rotate(Box(8, .2, .5), -.8)
        .Select(p => P(p.X + 7212103, p.Y + 1041196, p.Z + 48)).ToList(), new[] { Math.Cos(-.8), Math.Sin(-.8), 0 })),
    ("surface tessellation invariant", () => {
        var original = Rotate(Box(8, .2, .5), -.8, .4);
        var split = new List<Point3Dto>();
        for (int i = 0; i < original.Count; i += 3) {
            var a = original[i]; var b = original[i + 1]; var c = original[i + 2];
            if (i % 2 != 0) { split.AddRange(new[] { a, b, c }); continue; }
            var m = P((a.X + b.X + c.X) / 3, (a.Y + b.Y + c.Y) / 3, (a.Z + b.Z + c.Z) / 3);
            split.AddRange(new[] { a, b, m, b, c, m, c, a, m });
        }
        var x = Fit(original); var y = Fit(split);
        Assert(Math.Abs(x.Reliability - y.Reliability) < 1e-7, "density-biased covariance");
        Assert(Math.Abs(x.Axis.X * y.Axis.X + x.Axis.Y * y.Axis.Y + x.Axis.Z * y.Axis.Z) > .99999999, "density-biased axis");
    }),
    ("I section", () => {
        var beam = Box(8, .3, .03, -.235);
        beam.AddRange(Box(8, .3, .03, .235)); beam.AddRange(Box(8, .02, .44));
        CheckAxis(Rotate(beam, -1.1), new[] { Math.Cos(-1.1), Math.Sin(-1.1), 0 });
    }),
    ("channel section", () => {
        var beam = Box(8, .07, .008, -.076);
        beam.AddRange(Box(8, .07, .008, .076));
        beam.AddRange(Box(8, .008, .144).Select(p => P(p.X, p.Y - .031, p.Z)));
        CheckAxis(Rotate(beam, -1.168), new[] { Math.Cos(-1.168), Math.Sin(-1.168), 0 });
    }),
    ("cube ambiguous", () => Assert(!Fit(Box(1, 1, 1)).Usable, "cube accepted")),
    ("short plate ambiguous", () => Assert(!Fit(Box(.3, .2, .01)).Usable, "plate accepted")),
    ("empty", () => Assert(!Fit(new()).Usable, "empty accepted")),
    ("degenerate", () => Assert(!Fit(Enumerable.Repeat(P(1, 1, 1), 6).ToList()).Usable, "degenerate accepted")),
    ("incomplete triangles", () => Assert(!Fit(Box(8, .2, .5).Skip(1).ToList()).Usable, "partial accepted")),
    ("nonfinite", () => { var mesh = Box(8, .2, .5); mesh[0] = P(double.NaN, 0, 0); Assert(!Fit(mesh).Usable, "NaN accepted"); }),
};
foreach (var (name, test) in tests) { test(); Console.WriteLine("PASS " + name); }
Console.WriteLine($"{tests.Length} geometry tests passed.");

static AxisGeometryDto Fit(List<Point3Dto> mesh) => MeshAxisFitter.Fit(mesh, "mesh-surface-pca-row");
static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
static void ExpectFailure(Action action) {
    try { action(); } catch (InvalidOperationException) { return; }
    throw new Exception("Expected invalid instance failure");
}
static Point3Dto P(double x, double y, double z) => new Point3Dto { X = x, Y = y, Z = z };
static BoundDto Bounds(List<Point3Dto> p) => new BoundDto {
    Min = P(p.Min(v => v.X), p.Min(v => v.Y), p.Min(v => v.Z)),
    Max = P(p.Max(v => v.X), p.Max(v => v.Y), p.Max(v => v.Z))
};
static void CheckAxis(List<Point3Dto> mesh, double[] direction) {
    var result = Fit(mesh);
    Assert(result.Usable, result.Reason ?? "unusable");
    var dot = result.Axis.X * direction[0] + result.Axis.Y * direction[1] + result.Axis.Z * direction[2];
    Assert(Math.Abs(dot) > .999999, "wrong axis: " + dot);
    Assert(Math.Abs(result.Length - 8) < 1e-6, "wrong mesh length");
}
static List<Point3Dto> Rotate(List<Point3Dto> mesh, double azimuth, double elevation = 0) => mesh.Select(p => {
    var x = p.X * Math.Cos(elevation) - p.Z * Math.Sin(elevation);
    var z = p.X * Math.Sin(elevation) + p.Z * Math.Cos(elevation);
    return P(x * Math.Cos(azimuth) - p.Y * Math.Sin(azimuth), x * Math.Sin(azimuth) + p.Y * Math.Cos(azimuth), z);
}).ToList();
static List<Point3Dto> Box(double x, double y, double z, double zOffset = 0) {
    var p = new List<Point3Dto>();
    for (int i = 0; i < 8; i++) p.Add(P(((i & 1) == 0 ? -.5 : .5) * x,
        ((i & 2) == 0 ? -.5 : .5) * y, ((i & 4) == 0 ? -.5 : .5) * z + zOffset));
    int[] indices = { 0,2,3, 0,3,1, 4,5,7, 4,7,6, 0,1,5, 0,5,4, 2,6,7, 2,7,3, 0,4,6, 0,6,2, 1,3,7, 1,7,5 };
    return indices.Select(i => p[i]).ToList();
}

namespace KPLN_NavisMcpBridge {
    public class Point3Dto { public double X; public double Y; public double Z; }
    public class BoundDto { public Point3Dto Min; public Point3Dto Max; }
}
