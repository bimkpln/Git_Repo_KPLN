import importlib.util
import math
from pathlib import Path
import unittest


spec = importlib.util.spec_from_file_location(
    "navis_ceiling_service_test",
    Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py",
)
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


def row():
    return {
        "Name": "clash",
        "Item1Path": "КП_Краска / Многослойный потолок / Потолки / этаж / AR.nwc",
        "Item2Path": "L5_Ladder tray / Кабельные лотки / этаж / SS.nwc",
        "Item2AxisGeometry": {
            "Usable": True,
            "Reason": None,
            "Source": "mesh-surface-pca-row",
            "Axis": {"X": 0, "Y": 0, "Z": 1},
            "Length": 20,
            "TransverseSize": 0.5,
            "Reliability": 40,
        },
        "Item1CeilingPlane": {
            "Usable": True,
            "Reason": None,
            "Source": "ceiling-mesh-plane-1",
            "Normal": {"X": 0, "Y": 0, "Z": -1},
            "DominantAreaRatio": 50,
        },
    }


class CeilingServiceGeometryTests(unittest.TestCase):
    def test_tray_and_ceiling_are_classified_from_exact_categories(self):
        metrics = server._metrics(row())
        self.assertEqual(metrics["Class1"], "потолок")
        self.assertEqual(metrics["Class2"], "лоток")
        self.assertEqual(metrics["CeilingServiceMetricsVersion"], "ceiling-service-mesh-1")
        self.assertEqual(metrics["MepAxis"], {"X": 0.0, "Y": 0.0, "Z": 1.0})
        self.assertEqual(metrics["CeilingNormal"], {"X": 0.0, "Y": 0.0, "Z": -1.0})
        self.assertTrue(metrics["CeilingDirectionGeomOk"])

    def test_swapped_sides(self):
        result = row()
        for suffix in ("Path", "AxisGeometry", "CeilingPlane"):
            a, b = f"Item1{suffix}", f"Item2{suffix}"
            result[a], result[b] = result.get(b), result.get(a)
        metrics = server._metrics(result)
        self.assertTrue(metrics["CeilingDirectionGeomOk"])
        self.assertEqual(metrics["MepAxis"]["Z"], 1)
        self.assertEqual(metrics["CeilingNormal"]["Z"], -1)

    def test_inclined_unit_vectors_are_preserved(self):
        result = row()
        s = math.sqrt(0.5)
        result["Item2AxisGeometry"]["Axis"] = {"X": s, "Y": 0, "Z": s}
        result["Item1CeilingPlane"]["Normal"] = {"X": 0, "Y": s, "Z": s}
        metrics = server._metrics(result)
        self.assertTrue(metrics["CeilingDirectionGeomOk"])
        self.assertAlmostEqual(metrics["MepAxis"]["X"], s)
        self.assertAlmostEqual(metrics["CeilingNormal"]["Y"], s)

    def test_invalid_axis_or_plane_never_falls_back_to_aabb(self):
        for target, changes in (
            ("Item2AxisGeometry", {"Source": "aabb"}),
            ("Item2AxisGeometry", {"Reliability": 1}),
            ("Item2AxisGeometry", {"Axis": {"X": 0, "Y": 0, "Z": 2}}),
            ("Item1CeilingPlane", {"Source": "aabb"}),
            ("Item1CeilingPlane", {"DominantAreaRatio": 1.9}),
            ("Item1CeilingPlane", {"Normal": {"X": 0, "Y": 0, "Z": 2}}),
        ):
            with self.subTest(target=target, changes=changes):
                result = row()
                result[target].update(changes)
                metrics = server._metrics(result)
                self.assertFalse(metrics["CeilingDirectionGeomOk"])
                if target == "Item2AxisGeometry":
                    self.assertIsNone(metrics["MepAxis"])
                else:
                    self.assertIsNone(metrics["CeilingNormal"])

    def test_non_ceiling_pair_does_not_receive_direction_contract(self):
        result = row()
        result["Item1Path"] = "Базовая стена / Стены / этаж / AR.nwc"
        self.assertNotIn("CeilingServiceMetricsVersion", server._metrics(result))


if __name__ == "__main__":
    unittest.main()
