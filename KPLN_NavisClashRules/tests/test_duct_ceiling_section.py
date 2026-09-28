import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ceiling_services import classify_ceiling


class DuctSectionRuleTests(unittest.TestCase):
    def test_matching_section_overrides_longitudinal_estimate(self):
        row = dict(Class2="воздуховод", MepAxis=[1,0,0], CeilingNormal=[0,0,1],
                   DuctSectionFootprintMatch=True, DuctSectionCheckSource="object-section-properties+duct-ceiling-mesh-section")
        self.assertEqual(classify_ceiling(row,"OV"), ("Approved","transverse_duct_section_matches_ceiling_footprint"))

    def test_verified_transverse_short_duct_does_not_need_longest_axis(self):
        row = dict(Class2="воздуховод", DuctSectionFootprintMatch=True,
                   DuctSectionCheckSource="object-section-properties+duct-ceiling-mesh-section")
        self.assertEqual(classify_ceiling(row, "OV")[0], "Approved")

    def test_transverse_pipe_rule_is_not_changed_by_duct_feedback(self):
        row = dict(Class2="труба", MepAxis=[0,0,1], CeilingNormal=[0,0,1])
        self.assertEqual(classify_ceiling(row, "VK")[0], "Uncertain")

    def test_missing_section_check_does_not_approve_short_duct(self):
        row = dict(Item2Path="прямой / Воздуховоды / этаж / OV.nwc", MepAxis=[1,0,0],CeilingNormal=[0,0,1])
        self.assertEqual(classify_ceiling(row,"OV"), ("Uncertain","missing:duct_section_footprint"))

    def test_mismatch_needs_axis(self):
        row = dict(Class2="воздуховод", DuctSectionFootprintMatch=False,
                   DuctSectionCheckSource="object-section-properties+duct-ceiling-mesh-section",CeilingNormal=[0,0,1])
        self.assertEqual(classify_ceiling(row,"OV")[0],"Uncertain")
        row["MepAxis"]=[1,0,0]
        self.assertEqual(classify_ceiling(row,"OV")[0],"Approved")

    def test_aabb_or_truthy_string_not_matching_evidence(self):
        for match,source in ((True,"aabb"),("true","object-section-properties+duct-ceiling-mesh-section")):
            row = dict(Class2="воздуховод",MepAxis=[1,0,0],CeilingNormal=[0,0,1],
                       DuctSectionFootprintMatch=match,DuctSectionCheckSource=source)
            self.assertEqual(classify_ceiling(row,"OV")[0],"Uncertain")

    def test_other_services_keep_existing_rule(self):
        self.assertEqual(classify_ceiling(dict(Class2="труба",MepAxis=[1,0,0],CeilingNormal=[0,0,1]),"VK")[0],"Approved")

    def test_verified_duct_insulation_uses_same_approval_rule(self):
        row = dict(Class2="изоляция воздуховода", DuctSectionFootprintMatch=True,
                   DuctSectionCheckSource="object-section-properties+duct-ceiling-mesh-section")
        self.assertEqual(classify_ceiling(row, "OV"),
                         ("Approved", "transverse_duct_section_matches_ceiling_footprint"))

if __name__ == "__main__": unittest.main()
