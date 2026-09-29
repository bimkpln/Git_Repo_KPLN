import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from classify_opening_clashes import classify
from provenance import provenance


def pipe(**extra):
    row = {"Name": "sample", "PairClass": "стена+труба", "SectionMax": 100,
           "GeomOk": True, "WallNormalAngle": 0, "WallEndDistance": 1000,
           "WallEndFaceDistance": 1000, "WallPlanThickness": 200}
    row.update(extra)
    return row


class ClassificationTests(unittest.TestCase):
    def test_duct_insulation_follows_duct_threshold_using_outer_size(self):
        for size, expected in ((126, "Approved"), (150, "Approved"), (176, "Active"), (526, "Active")):
            row = {"Name": "insulation", "PairClass": "изоляция воздуховода+стена",
                   "SectionSource": "Object/Duct Size", "SectionMin": 126,
                   "SectionMax": size, "InsulationThicknessMm": 13}
            self.assertEqual(classify([row], 150)["insulation"][0], expected)
            for key in ("SectionSource", "InsulationThicknessMm", "SectionMax"):
                incomplete = dict(row, **{key: None})
                self.assertEqual(classify([incomplete], 150)["insulation"][0], "Uncertain")

    def test_duct_and_insulation_are_not_two_bundle_members(self):
        rows = [{"Name": "duct", "Group": "example", "PairClass": "воздуховод+стена",
                 "SectionMin": 100, "SectionMax": 100},
                {"Name": "shell", "Group": "example", "PairClass": "изоляция воздуховода+стена",
                 "SectionMin": 126, "SectionMax": 126, "SectionSource": "Object/Duct Size",
                 "InsulationThicknessMm": 13}]
        self.assertTrue(all(value[0] == "Approved" for value in classify(rows, 150).values()))
        rows[1]["SectionMax"] = 176
        self.assertEqual(classify(rows, 150)["shell"][0], "Active")

    def test_multiple_unassociated_insulation_members_remain_uncertain(self):
        rows = [{"Name": name, "Group": "example", "PairClass": "изоляция воздуховода+стена",
                 "SectionMin": 126, "SectionMax": 126, "SectionSource": "Object/Duct Size",
                 "InsulationThicknessMm": 13} for name in ("a", "b")]
        self.assertTrue(all(value[0] == "Uncertain" for value in classify(rows, 150).values()))

    def test_two_insulated_ducts_are_a_bundle_without_counting_shells_twice(self):
        rows = []
        for i in range(2):
            point = dict(X=i*100/304.8, Y=i*100/304.8, Z=0)
            bound = {"Min": point, "Max": {k: v+100/304.8 for k, v in point.items()}}
            rows.append(dict(Name=f"duct{i}", Group="example", PairClass="воздуховод+стена",
                             SectionMin=100, SectionMax=100, Item2Bound=bound, Bound=bound))
            outer = {"Min": {k: v-13/304.8 for k,v in bound["Min"].items()},
                     "Max": {k: v+13/304.8 for k,v in bound["Max"].items()}}
            rows.append(dict(Name=f"shell{i}", Group="example", PairClass="изоляция воздуховода+стена",
                             SectionMin=126, SectionMax=126, Item2Bound=outer, Bound=bound,
                             SectionSource="Object/Duct Size", InsulationThicknessMm=13))
        self.assertTrue(all(value[0] == "Active" for value in classify(rows, 150).values()))
        self.assertTrue(all(value[0] == "Active" for value in classify(rows, 250).values()))
        self.assertTrue(all(value[0] == "Approved" for value in classify(rows, 300).values()))

    def test_ceiling_scope_all_five_disciplines_and_no_size_threshold(self):
        for discipline in ("OV", "VK", "PT", "EOM", "SS"):
            row = {"Name": "sample", "MepAxis": [10, 0, 0], "CeilingNormal": [0, 0, -2],
                   "SectionMax": 900}
            self.assertEqual(classify([row], mode="ceiling-services", discipline=discipline)["sample"][0], "Approved")

    def test_ceiling_transverse_boundary_and_missing_vectors_do_not_approve(self):
        for axis in ([0, 0, 1], [1, 0, 1], [0, 0, 0], None):
            row = {"Name": "sample", "MepAxis": axis, "CeilingNormal": [0, 0, 1]}
            self.assertEqual(classify([row], mode="ceiling-services", discipline="PT")["sample"][0], "Uncertain")

    def test_sloped_ceiling_uses_actual_normal(self):
        row = {"Name": "sample", "MepAxis": [1, 0, 1], "CeilingNormal": [1, 0, -1]}
        self.assertEqual(classify([row], mode="ceiling-services", discipline="SS")["sample"][0], "Approved")
        row["CeilingNormal"] = [1, 0, 1]
        self.assertEqual(classify([row], mode="ceiling-services", discipline="SS")["sample"][0], "Uncertain")

    def test_threshold_is_strict_and_geometry_precedes_size(self):
        self.assertEqual(classify([pipe(SectionMax=150)], 150)["sample"][0], "Approved")
        self.assertEqual(classify([pipe(SectionMax=151)], 150)["sample"][0], "Active")
        self.assertEqual(classify([pipe(WallEndFaceDistance=0)], 150)["sample"][0], "Active")
        self.assertEqual(classify([pipe(WallNormalAngle=60)], 150)["sample"][0], "Active")

    def test_insulation_uses_outer_size(self):
        row = pipe(PairClass="изоляция трубы+стена", SectionMax=25, MepMinEdge=180)
        self.assertEqual(classify([row], 150)["sample"][0], "Active")

    def test_single_pipe_uses_own_section_when_wall_geometry_is_incomplete(self):
        for row in (pipe(GeomOk=False, WallNormalAngle=None, WallEndDistance=None,
                         WallEndFaceDistance=None, WallPlanThickness=None),
                    pipe(WallEndFaceDistance=None)):
            self.assertEqual(classify([row], 150)["sample"],
                             ("Approved", "single_pipe_section_in_tolerance"))
        self.assertEqual(classify([pipe(SectionMax=151, GeomOk=False)], 150)["sample"][0], "Active")
        self.assertEqual(classify([pipe(WallNormalAngle=60, WallEndFaceDistance=None)], 150)["sample"],
                         ("Active", "pipe_runs_longitudinally_in_wall"))

    def test_multi_pipe_group_still_requires_wall_geometry(self):
        rows = [pipe(Name=name, Group="example", GeomOk=False,
                     WallNormalAngle=None, WallEndDistance=None,
                     WallEndFaceDistance=None, WallPlanThickness=None)
                for name in ("a", "b")]
        self.assertTrue(all(value[0] == "Uncertain" for value in classify(rows, 150).values()))

    def test_pipe_against_finish_is_supported(self):
        row = pipe(PairClass="отделка+труба", GeomOk=False)
        self.assertEqual(classify([row], 150)["sample"],
                         ("Approved", "finish_layer_no_separate_opening"))

    def test_unknown_scope_and_missing_size_are_uncertain(self):
        for row in (pipe(PairClass="перекрытие+труба"), pipe(SectionMax=None),
                    pipe(SectionMax=float("nan"))):
            self.assertEqual(classify([row], 150)["sample"][0], "Uncertain")

    def test_duplicate_names_and_invalid_threshold_rejected(self):
        with self.assertRaises(ValueError):
            classify([pipe(), pipe()], 150)
        for value in (0, -1, float("inf")):
            with self.assertRaises(ValueError):
                classify([pipe()], value)

    def test_mixed_wall_group_does_not_propagate_approval(self):
        rows = [pipe(Name="a", Group="example"), pipe(Name="b", Group="example", WallEndDistance=0)]
        self.assertEqual([classify(rows, 150)[name][0] for name in ("a", "b")], ["Approved", "Active"])

    def test_duct_bundle_requires_complete_evidence(self):
        rows = [{"Name": name, "Group": "example", "PairClass": "воздуховод+стена",
                 "SectionMax": 100, "SectionMin": 80} for name in ("a", "b")]
        self.assertTrue(all(value[0] == "Uncertain" for value in classify(rows, 150).values()))
        for index, row in enumerate(rows):
            point = {"X": index * 100 / 304.8, "Y": index * 100 / 304.8, "Z": 0}
            row["Bound"] = {"Min": point, "Max": point}
        self.assertTrue(all(value[0] == "Active" for value in classify(rows, 150).values()))

    def test_cli_includes_decisions_and_provenance(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "input.json"
            path.write_text(json.dumps([pipe()]), encoding="utf-8-sig")
            result = subprocess.run([sys.executable, "-B", str(ROOT / "src/classify_opening_clashes.py"),
                                     str(path), "--opening-min-edge-mm", "150", "--compare-status"],
                                    check=True, capture_output=True, text=True, encoding="utf-8")
        output = json.loads(result.stdout)
        self.assertEqual(output["results"][0]["decision"], "Approved")
        self.assertEqual(len(output["provenance"]["ruleset_sha256"]), 64)
        self.assertEqual(output["provenance"]["parameters"]["opening_min_edge_mm"], 150)

    def test_fingerprint_tracks_policy_changes_but_not_line_endings(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "VERSION").write_bytes(b"0.1.0\n")
            (root / "rules").mkdir()
            policy = root / "rules/policy.md"
            policy.write_bytes(b"sample\n")
            a = provenance([pipe()], 150, root)
            policy.write_bytes(b"sample\r\n")
            self.assertEqual(a["ruleset_sha256"], provenance([pipe()], 150, root)["ruleset_sha256"])
            policy.write_bytes(b"changed\n")
            self.assertNotEqual(a["ruleset_sha256"], provenance([pipe()], 150, root)["ruleset_sha256"])
            self.assertNotEqual(a["input_sha256"], provenance([pipe(SectionMax=120)], 150, root)["input_sha256"])


if __name__ == "__main__":
    unittest.main()
