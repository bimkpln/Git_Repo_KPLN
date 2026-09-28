import importlib.util
import math
from pathlib import Path
import unittest
from unittest.mock import patch

import httpx

spec = importlib.util.spec_from_file_location(
    "navis_frame_test", Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py")
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


def row():
    s = math.sqrt(.5)
    return {
        "Name": "clash", "Status": "eTestResultStatus_APPROVED",
        "Item1Path": "Базовая стена / Стены / этаж / AR.nwc",
        "Item2Path": "С255 / Швеллер (НесКаркас_Балка) / Каркас несущий / этаж / KR.nwc",
        "Item1Bound": {"Min": dict(X=0, Y=0, Z=0), "Max": dict(X=4, Y=4, Z=4)},
        "Item2Bound": {"Min": dict(X=0, Y=0, Z=0), "Max": dict(X=4, Y=4, Z=.5)},
        "Item1PlanGeometry": dict(AxisX=s, AxisY=-s, CenterX=2, CenterY=2,
                                  HalfLength=3, HalfThickness=.1, Reliability=20,
                                  Source="mesh-pca-row"),
        "Item2AxisGeometry": dict(Axis=dict(X=s, Y=-s, Z=0), Length=8,
                                  TransverseSize=.5, Reliability=30,
                                  Source="mesh-surface-pca-row", Usable=True, Reason=None),
    }


class FrameMetricsTests(unittest.TestCase):
    def test_parallel_negative_xy(self):
        m = server._metrics(row())
        self.assertEqual(m["MetricsVersion"], "duct-footprint-6")
        self.assertTrue(m["GeomOk"])
        self.assertAlmostEqual(m["FrameWallNormalAngle"], 90)
        self.assertEqual(m["WallNormalAngle"], m["FrameWallNormalAngle"])
        self.assertEqual(m["FrameLengthMm"], 2438.4)
        for key in ("Angle", "PerpRel", "Dia2", "Len2", "SectionMax", "Through"):
            self.assertIsNone(m[key])

    def test_same_bounds_perpendicular_mesh(self):
        r = row()
        r["Item2AxisGeometry"]["Axis"]["Y"] *= -1
        self.assertAlmostEqual(server._metrics(r)["FrameWallNormalAngle"], 0)

    def test_swapped_sides(self):
        r = row()
        for suffix in ("Path", "Bound", "PlanGeometry", "AxisGeometry"):
            a, b = f"Item1{suffix}", f"Item2{suffix}"
            r[a], r[b] = r.get(b), r.get(a)
        self.assertAlmostEqual(server._metrics(r)["FrameWallNormalAngle"], 90)

    def test_old_bridge_no_aabb_fallback(self):
        r = row()
        del r["Item2AxisGeometry"]
        m = server._metrics(r)
        self.assertFalse(m["GeomOk"])
        self.assertIsNone(m["WallNormalAngle"])
        self.assertIsNone(m["Angle"])

    def test_bad_mesh(self):
        for changes in ({"Usable": False}, {"Reason": "mesh-incomplete-or-point-limit"},
                        {"Source": "aabb"}, {"Reliability": 1}, {"Length": .5},
                        {"Axis": dict(X=float("nan"), Y=0, Z=0)},
                        {"Axis": dict(X=2, Y=0, Z=0)}, {"Axis": None}):
            with self.subTest(changes=changes):
                r = row()
                r["Item2AxisGeometry"].update(changes)
                m = server._metrics(r)
                self.assertFalse(m["GeomOk"])
                self.assertIsNone(m["WallNormalAngle"])

    def test_wall_unreliable(self):
        for changes in ({"Reliability": 1}, {"AxisX": float("nan")}, {"Source": "aabb"}, {"AxisY": 6}):
            r = row()
            r["Item1PlanGeometry"].update(changes)
            self.assertFalse(server._metrics(r)["GeomOk"])

    def test_vertical_and_inclined_axes(self):
        r = row()
        r["Item2AxisGeometry"]["Axis"] = dict(X=0, Y=0, Z=1)
        self.assertAlmostEqual(server._metrics(r)["FrameWallNormalAngle"], 90)
        r["Item2AxisGeometry"]["Axis"] = dict(X=.5, Y=.5, Z=math.sqrt(.5))
        self.assertAlmostEqual(server._metrics(r)["FrameWallNormalAngle"], 45)

    def test_wall_finish_uses_wall_mesh(self):
        r = row()
        r["Item1Path"] = "штукатурка / " + r["Item1Path"]
        self.assertEqual(server._metrics(r)["Class1"], "отделка")
        self.assertTrue(server._metrics(r)["GeomOk"])

    def test_frame_frame_has_no_pipe_metrics(self):
        r = row()
        r["Item1Path"] = r["Item2Path"]
        self.assertFalse(server._metrics(r)["GeomOk"])
        self.assertIsNone(server._metrics(r)["Angle"])

    def test_transport_retains_axis_when_bounds_removed(self):
        requests = []
        def handler(request):
            requests.append(request)
            r = row()
            r["GeometryVersion"] = "instance-mesh-2"
            r["Item2AxisGeometry"]["Diagnostics"] = dict(InstancePath="1/2/3", MatchedFragments=1,
                SkippedInstanceFragments=999, GeneratedTriangleCount=285)
            return httpx.Response(200, json=[r])
        client = httpx.Client(base_url=server.BRIDGE_URL, transport=httpx.MockTransport(handler))
        self.addCleanup(client.close)
        with patch.object(server, "_client", return_value=client):
            results = server.get_clash_test_results("test", status="Approved", include_metrics=True,
                                                    include_geometry=False)
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0].method, "GET")
        self.assertEqual(requests[0].url.params["status"], "Approved")
        self.assertEqual(requests[0].url.params["includeItemBounds"], "true")
        self.assertIn("Item2AxisGeometry", results[0])
        self.assertEqual(results[0]["GeometryVersion"], "instance-mesh-2")
        self.assertEqual(results[0]["MetricsVersion"], "duct-footprint-6")
        self.assertEqual(results[0]["Item2AxisGeometry"]["Diagnostics"]["SkippedInstanceFragments"], 999)
        self.assertNotIn("Item2Bound", results[0])
        self.assertEqual(results[0]["Status"], "eTestResultStatus_APPROVED")

    def test_pipe_metrics_unchanged(self):
        for path in ("Труба / Трубы / этаж / VK.nwc", "Воздуховод / Воздуховоды / этаж / OV.nwc"):
            r = row()
            r["Item2Path"] = path
            del r["Item2AxisGeometry"]
            m = server._metrics(r)
            self.assertFalse(m["GeomOk"])  # No clash Bound / WallEndDistance in this fixture.
            self.assertAlmostEqual(m["WallNormalAngle"], 0)
            self.assertEqual(m["SectionMax"], 152.4)
            self.assertEqual(m["MepMinEdge"], 152.4)
            self.assertNotIn("FrameAxis", m)


if __name__ == "__main__":
    unittest.main()
