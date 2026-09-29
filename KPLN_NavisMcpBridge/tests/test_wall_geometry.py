import importlib.util
import math
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location(
    "navis_wall_test", Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py")
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


def row():
    s = math.sqrt(0.5)
    return {
        "Item1Path": "Гипс / Базовая стена / Стены / AR.nwc",
        "Item2Path": "Сталь / Трубы / OV.nwc",
        "Item1Bound": {"Min": dict(X=-3, Y=-3, Z=0), "Max": dict(X=3, Y=3, Z=10)},
        "Item2Bound": {"Min": dict(X=-2, Y=-2, Z=4.9), "Max": dict(X=2, Y=2, Z=5.1)},
        "Item2Size": "60 мм",
        "WallGeometryVersion": "wall-instance-mesh-1",
        "Pt1": dict(X=0, Y=0, Z=5),
        "Bound": {"Min": dict(X=40, Y=40, Z=0), "Max": dict(X=50, Y=50, Z=10)},
        "Item1PlanGeometry": dict(Source="wall-instance-mesh-1", Usable=True, Reason=None,
            AxisX=s, AxisY=-s, CenterX=0, CenterY=0, HalfLength=4, HalfThickness=.1,
            Reliability=20, NearestBroadFace=0, NearestEndFace=4),
        "Item2AxisGeometry": dict(Source="mesh-surface-pca-row", Usable=True, Reason=None,
            Axis=dict(X=s, Y=-s, Z=0), Reliability=50, Length=6, TransverseSize=.2),
    }


class WallGeometryTests(unittest.TestCase):
    def test_signed_diagonal_axis_not_absolute_aabb(self):
        result = row()
        metrics = server._metrics(result)
        self.assertTrue(metrics["GeomOk"])
        self.assertEqual(metrics["WallNormalAngle"], 90)
        self.assertEqual(metrics["SectionMax"], 60)
        self.assertEqual(metrics["Kind2"], "голая труба")
        result["Item2AxisGeometry"]["Axis"]["Y"] *= -1
        self.assertEqual(server._metrics(result)["WallNormalAngle"], 0)

    def test_actual_contact_not_clash_union_center(self):
        result = row()
        result["Item1PlanGeometry"].update(AxisX=1, AxisY=0)
        metrics = server._metrics(result)
        self.assertAlmostEqual(metrics["WallEndDistance"], 1219.2)
        self.assertAlmostEqual(metrics["WallPlanThickness"], 61.0)

    def test_failure_reason_survives_without_fabricated_angle(self):
        result = row()
        result["Item1PlanGeometry"].update(Usable=False, Reason="wall-element-owner-not-found")
        metrics = server._metrics(result)
        self.assertFalse(metrics["GeomOk"])
        self.assertIsNone(metrics["WallNormalAngle"])
        self.assertIsNone(metrics["WallEndDistance"])
        self.assertEqual(metrics["WallGeometryReason"], "wall-element-owner-not-found")

    def test_legacy_single_face_never_proves_direction(self):
        result = row()
        result["Item1PlanGeometry"].update(Source="mesh-pca-row", HalfThickness=0, Reliability=1e12)
        metrics = server._metrics(result)
        self.assertFalse(metrics["GeomOk"])
        self.assertIsNone(metrics["WallNormalAngle"])

    def test_new_geometry_never_falls_back_to_unsigned_axis(self):
        for geometry in (None, dict(Usable=False, Reason="ambiguous-axis")):
            result = row()
            result["Item2AxisGeometry"] = geometry
            metrics = server._metrics(result)
            self.assertFalse(metrics["GeomOk"])
            self.assertIsNone(metrics["WallNormalAngle"])
            self.assertIsNotNone(metrics["WallMepAxisReason"])

    def test_wall_invalid_normals_and_thickness(self):
        for changes in (dict(AxisX=2), dict(HalfThickness=-1), dict(AxisX=float("nan")),
                        dict(Reliability=1), dict(Source="aabb")):
            with self.subTest(changes=changes):
                result = row()
                result["Item1PlanGeometry"].update(changes)
                self.assertIsNone(server._metrics(result)["WallNormalAngle"])

    def test_swapped_sides_and_pipe_insulation(self):
        result = row()
        result["Item2Path"] = "Изоляция / Материалы изоляции труб / OV.nwc"
        for suffix in ("Path", "Bound", "Size", "PlanGeometry", "AxisGeometry"):
            a, b = f"Item1{suffix}", f"Item2{suffix}"
            result[a], result[b] = result.get(b), result.get(a)
        metrics = server._metrics(result)
        self.assertEqual(metrics["WallNormalAngle"], 90)
        self.assertTrue(metrics["GeomOk"])


if __name__ == "__main__":
    unittest.main()
