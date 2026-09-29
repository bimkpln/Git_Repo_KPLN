import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from classify_opening_clashes import classify


def clash(kind1="изоляция", kind2="изоляция", **extra):
    row = {
        "Name": "sample",
        "Kind1": kind1,
        "Kind2": kind2,
        "GeomOk": True,
        "PerpRel": 0.5,
        "Angle": 90,
        "Elong1": 2,
        "Elong2": 2,
    }
    row.update(extra)
    return row


def decision(row):
    return classify([row], mode="self-intersections")["sample"]


class SelfIntersectionTests(unittest.TestCase):
    def test_insulation_or_fitting_against_bare_body_is_active_without_geometry(self):
        for insulated_kind in ("изоляция", "фитинг"):
            row = clash("голая труба", insulated_kind, GeomOk=False,
                        PerpRel=None, Angle=None, Elong1=None, Elong2=None)
            self.assertEqual(decision(row),
                             ("Active", "insulated_envelope_reaches_bare_body"))

    def test_bare_body_pair_is_not_automatically_classified(self):
        self.assertEqual(decision(clash("голая труба", "голая труба")),
                         ("Uncertain", "bare_body_pair_not_trained"))

    def test_axis_crossing_boundary_is_active(self):
        self.assertEqual(decision(clash(PerpRel=0.15))[0], "Active")
        self.assertEqual(decision(clash(PerpRel=0.150001))[0], "Approved")

    def test_parallel_extended_boundary_is_active(self):
        self.assertEqual(decision(clash(Angle=5, Elong1=5))[0], "Active")
        self.assertEqual(decision(clash(Angle=5.001, Elong1=5))[0], "Approved")
        self.assertEqual(decision(clash(Angle=5, Elong1=4.999, Elong2=4))[0], "Approved")

    def test_transverse_or_local_insulated_contact_is_approved(self):
        self.assertEqual(decision(clash()),
                         ("Approved", "transverse_or_local_insulated_contact"))
        self.assertEqual(decision(clash("изоляция", "фитинг"))[0], "Approved")

    def test_missing_geometry_needed_to_exclude_active_rules_is_uncertain(self):
        self.assertEqual(decision(clash(GeomOk=False))[0], "Uncertain")
        self.assertEqual(decision(clash(PerpRel=None))[0], "Uncertain")
        self.assertEqual(decision(clash(Angle=None))[0], "Uncertain")
        self.assertEqual(decision(clash(Angle=5, Elong2=None))[0], "Uncertain")

    def test_unknown_kind_is_uncertain(self):
        self.assertEqual(decision(clash("неизвестно", "изоляция"))[0], "Uncertain")


if __name__ == "__main__":
    unittest.main()
