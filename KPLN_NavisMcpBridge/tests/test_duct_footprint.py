import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location("duct_footprint_server", Path(__file__).resolve().parents[1] / "server" / "navis_mcp_server.py")
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


def row(a=188, b=473, area=None):
    return dict(Name="sample", Item1Path="материал / Потолки / этаж / AR.nwc",
                Item2Path="прямоугольный / Воздуховоды / этаж / OV.nwc",
                DuctCeilingFootprint=dict(Usable=True, Source="duct-ceiling-mesh-section", DuctSide=2,
                    ServiceKind="duct", NominalSizeSource="Object/Height+Width",
                    NominalWidth=188/304.8, NominalHeight=473/304.8,
                    MinEdge=a/304.8, MaxEdge=b/304.8, Area=(a*b if area is None else area)/304.8**2,
                    DuctSliceArea=a*b/304.8**2, Normal=dict(X=0,Y=0,Z=1)))


def round_row(diameter=100, area=None):
    circle_area = 3.141592653589793 * diameter**2 / 4
    r = row(diameter, diameter, circle_area if area is None else area)
    r["DuctCeilingFootprint"].update(ServiceKind="duct", NominalSizeSource="Object/Diameter",
        NominalWidth=None, NominalHeight=None, NominalDiameter=diameter/304.8,
        DuctSliceArea=circle_area/304.8**2)
    return r


def insulation_row(size="150x500", thickness=13, a=176, b=526, area=None):
    r = row(a, b, a*b if area is None else area)
    r["Item2Path"] = "изоляция / Материалы изоляции воздуховодов / этаж / OV.nwc"
    r["DuctCeilingFootprint"].update(ServiceKind="duct-insulation", NominalSizeSource="Object/Duct Size",
        NominalSize=size, NominalWidth=None, NominalHeight=None,
        InsulationThickness=thickness/304.8)
    return r


class DuctFootprintTests(unittest.TestCase):
    def test_insulation_requires_thickness(self):
        r = insulation_row()
        r["DuctCeilingFootprint"]["InsulationThickness"] = None
        self.assertIsNone(server._metrics(r)["DuctSectionFootprintMatch"])

    def test_wall_insulation_uses_exact_outer_section_on_either_side(self):
        for size, expected in (("150×500", (176, 526)), ("⌀100", (126, 126)),
                               ("ø100", (126, 126)), ("100х150", (126, 176))):
            for reverse in (False, True):
                r = insulation_row(size)
                section = r.pop("DuctCeilingFootprint")
                r.update(Item1Path="материал / Стены / этаж / KR.nwc",
                         Item2DuctSection=section, Item2Size="WRONG",
                         Item1Bound={"Min":dict(X=0,Y=0,Z=0), "Max":dict(X=10,Y=1,Z=10)},
                         Item2Bound={"Min":dict(X=2,Y=-1,Z=2), "Max":dict(X=3,Y=2,Z=3)})
                if reverse:
                    for suffix in ("Path", "Bound", "DuctSection", "Size"):
                        a, b = f"Item1{suffix}", f"Item2{suffix}"
                        r[a], r[b] = r.get(b), r.get(a)
                m = server._metrics(r)
                self.assertEqual((m["SectionMin"], m["SectionMax"]), expected)
                self.assertEqual(m["SectionSource"], "Object/Duct Size")

    def test_wall_insulation_does_not_fall_back_to_legacy_size_or_box(self):
        r = insulation_row()
        r.pop("DuctCeilingFootprint")
        r.update(Item1Path="Стены / этаж / KR.nwc", Item2Size="ø100",
                 Item1Bound={"Min":dict(X=0,Y=0,Z=0), "Max":dict(X=10,Y=1,Z=10)},
                 Item2Bound={"Min":dict(X=2,Y=-1,Z=2), "Max":dict(X=3,Y=2,Z=3)})
        self.assertIsNone(server._metrics(r)["SectionMax"])

    def test_wall_direct_duct_prefers_numeric_dimensions(self):
        r = row()
        section = r.pop("DuctCeilingFootprint")
        r.update(Item1Path="Стены / этаж / KR.nwc", Item2Size="999x999",
                 Item2DuctSection=section,
                 Item1Bound={"Min":dict(X=0,Y=0,Z=0), "Max":dict(X=10,Y=1,Z=10)},
                 Item2Bound={"Min":dict(X=2,Y=-1,Z=2), "Max":dict(X=3,Y=2,Z=3)})
        self.assertEqual(server._metrics(r)["SectionMax"], 473)

    def test_example(self):
        m=server._metrics(row())
        self.assertTrue(m["DuctSectionFootprintMatch"])
        self.assertEqual(m["DuctSectionAreaMm2"], 88924)
        self.assertEqual(m["DuctCeilingRelation"], "transverse-by-section")
        self.assertIsNone(m["Angle"])
        self.assertIsNone(m["MepAxis"])
        self.assertFalse(m["CeilingDirectionGeomOk"])

    def test_round_duct_uses_numeric_diameter_and_circle_area(self):
        m = server._metrics(round_row())
        self.assertTrue(m["DuctSectionFootprintMatch"])
        self.assertEqual(m["DuctSectionShape"], "round")
        self.assertAlmostEqual(m["DuctSectionAreaMm2"], 3.141592653589793*100**2/4)

    def test_rectangular_insulation_expands_duct_size_by_thickness(self):
        m = server._metrics(insulation_row())
        self.assertTrue(m["DuctSectionFootprintMatch"])
        self.assertEqual(m["SectionMin"], 176)
        self.assertEqual(m["SectionMax"], 526)
        self.assertEqual(m["InsulationThicknessMm"], 13)

    def test_round_insulation_parses_diameter_and_expands_by_thickness(self):
        diameter = 126
        area = 3.141592653589793 * diameter**2 / 4
        r = insulation_row("ø100", 13, diameter, diameter, area)
        r["DuctCeilingFootprint"]["DuctSliceArea"] = area/304.8**2
        m = server._metrics(r)
        self.assertTrue(m["DuctSectionFootprintMatch"])
        self.assertEqual(m["DuctSectionShape"], "round")

    def test_both_edges_not_only_large_edge(self):
        self.assertFalse(server._metrics(row(100,473))["DuctSectionFootprintMatch"])

    def test_same_area_different_edges(self):
        self.assertFalse(server._metrics(row(94,946))["DuctSectionFootprintMatch"])

    def test_same_edges_with_hole(self):
        self.assertFalse(server._metrics(row(area=88924-10000))["DuctSectionFootprintMatch"])

    def test_area_tolerance_is_two_percent(self):
        self.assertTrue(server._metrics(row(area=188*473*.981))["DuctSectionFootprintMatch"])
        self.assertFalse(server._metrics(row(area=188*473*.979))["DuctSectionFootprintMatch"])

    def test_partial_longitudinal_slice_cannot_mimic_nominal_section(self):
        r = row()
        r["DuctCeilingFootprint"]["DuctSliceArea"] = 188*2000/304.8**2
        m = server._metrics(r)
        self.assertFalse(m["FullDuctSectionCovered"])
        self.assertFalse(m["DuctSectionFootprintMatch"])

    def test_missing_slice_area_is_unknown(self):
        r = row()
        del r["DuctCeilingFootprint"]["DuctSliceArea"]
        self.assertIsNone(server._metrics(r)["DuctSectionFootprintMatch"])

    def test_small_measurement_difference(self):
        self.assertTrue(server._metrics(row(189,474))["DuctSectionFootprintMatch"])

    def test_not_matching_is_not_longitudinal(self):
        m=server._metrics(row(188,2000))
        self.assertFalse(m["DuctSectionFootprintMatch"])
        self.assertEqual(m["DuctCeilingRelation"],"unconfirmed")

    def test_insulation_size_parser_handles_swapped_sides_and_decimal_comma(self):
        for size in ("473x188","188,0х473,0", "473 mm × 188 mm", "188x473-188x473"):
            self.assertTrue(server._metrics(insulation_row(size,13,214,499))["DuctSectionFootprintMatch"])

    def test_invalid_or_transition_insulation_size(self):
        for size in (None,"", "188x473-200x500", "0x473", "188x473 м", "ø100-ø125"):
            self.assertIsNone(server._metrics(insulation_row(size,13,214,499))["DuctSectionFootprintMatch"])

    def test_direct_duct_requires_dedicated_dimension_properties(self):
        r=row();r["DuctCeilingFootprint"].update(NominalSize="188x473",NominalSizeSource="Object/Size")
        self.assertIsNone(server._metrics(r)["DuctSectionFootprintMatch"])

    def test_missing_mesh_no_bbox_fallback(self):
        r=row();del r["DuctCeilingFootprint"]
        r["Bound"]={"Min":dict(X=0,Y=0,Z=0),"Max":dict(X=188/304.8,Y=473/304.8,Z=.1)}
        self.assertIsNone(server._metrics(r)["DuctSectionFootprintMatch"])

    def test_bad_mesh_rejected(self):
        for changes in ({"Usable":False},{"Area":float("nan")},{"Area":-1},{"Source":"aabb"},{"DuctSide":1},{"Normal":dict(X=0,Y=0,Z=2)}):
            r=row();r["DuctCeilingFootprint"].update(changes)
            self.assertIsNone(server._metrics(r)["DuctSectionFootprintMatch"])

    def test_swapped_participants(self):
        r=row();r["Item1Path"],r["Item2Path"]=r["Item2Path"],r["Item1Path"]
        r["DuctCeilingFootprint"]["DuctSide"]=1
        self.assertTrue(server._metrics(r)["DuctSectionFootprintMatch"])

if __name__ == "__main__": unittest.main()
