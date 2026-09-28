import math
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from classify_opening_clashes import classify
from slab_openings import polygon_gap


def pipe(name="a", x=0, diameter=100, group=None, **extra):
    points = [[x+diameter/2*math.cos(i*math.pi/16), diameter/2*math.sin(i*math.pi/16), 0] for i in range(32)]
    row = dict(Name=name, Status="Active", PairClass="перекрытие+труба", SectionMax=diameter,
               GeomOk=True, SlabMetricsVersion="pipe-slab-mesh-1", SlabPipeAxis=[0,0,1],
               SlabNormal=[0,0,1], SlabNormalAngle=0, SlabOuterDiameterMm=diameter,
               SlabEdgeContact=False, SlabFullSectionCovered=True, PipeId=name, SlabId="slab",
               GroupId=group, GroupIdentityKnown=True, Group="same display name",
               SlabSections=[dict(origin_mm=[0,0,0], contour_mm=points)])
    row.update(extra)
    return row


def decisions(rows, threshold=150):
    return {key: value[0] for key, value in classify(rows, threshold, "slab-openings").items()}


class SlabTests(unittest.TestCase):
    def test_nominal_threshold_inclusive(self):
        for size, expected in ((100,"Approved"),(150,"Approved"),(150.01,"Active")):
            self.assertEqual(decisions([pipe(diameter=size)])["a"], expected)

    def test_geometry_before_size(self):
        self.assertEqual(decisions([pipe(SlabEdgeContact=True)])["a"], "Active")
        self.assertEqual(decisions([pipe(SlabPipeAxis=[1,0,0], SlabNormalAngle=90)])["a"], "Active")

    def test_missing_geometry_preserves_uncertainty(self):
        for key in ("SlabPipeAxis", "SlabNormal", "SlabNormalAngle", "PipeId", "SlabId", "SlabOuterDiameterMm"):
            self.assertEqual(decisions([pipe(**{key:None})])["a"], "Uncertain")
        self.assertEqual(decisions([pipe(SlabFullSectionCovered=False, SlabSections=[])])["a"], "Uncertain")
        self.assertEqual(decisions([pipe(GroupIdentityKnown=False)])["a"], "Uncertain")

    def test_45_degree_boundary(self):
        s=math.sqrt(.5)
        self.assertEqual(decisions([pipe(SlabPipeAxis=[s,0,s], SlabNormalAngle=45)])["a"],"Uncertain")

    def test_bundle_gap_strictly_less_than_larger_diameter(self):
        self.assertEqual(decisions([pipe(group="g"),pipe("b",200,group="g")]),{"a":"Approved","b":"Approved"})
        self.assertEqual(decisions([pipe(group="g"),pipe("b",199.9,group="g")]),{"a":"Active","b":"Active"})

    def test_bundle_uses_actual_gap_not_center_distance(self):
        self.assertEqual(decisions([pipe(group="g"),pipe("b",160,group="g")]),{"a":"Active","b":"Active"})

    def test_bundle_uses_larger_member(self):
        self.assertEqual(decisions([pipe(group="g"),pipe("b",170,diameter=60,group="g")]),{"a":"Active","b":"Active"})

    def test_bundle_below_threshold_and_chain(self):
        self.assertTrue(all(x=="Approved" for x in decisions([pipe(diameter=65,group="g"),pipe("b",80,65,"g")]).values()))
        self.assertTrue(all(x=="Active" for x in decisions([pipe(group="g"),pipe("b",160,group="g"),pipe("c",320,group="g")]).values()))

    def test_same_names_but_different_group_ids(self):
        self.assertTrue(all(x=="Approved" for x in decisions([pipe(group="g1"),pipe("b",160,group="g2")]).values()))

    def test_different_slabs_not_combined(self):
        self.assertTrue(all(x=="Approved" for x in decisions([pipe(group="g"),pipe("b",160,group="g",SlabId="other")]).values()))

    def test_same_pipe_not_counted_twice(self):
        self.assertTrue(all(x=="Approved" for x in decisions([pipe(group="g"),pipe("b",group="g",PipeId="a")]).values()))

    def test_workflow_exclusion_before_bundle(self):
        for status in ("Resolved","Reviewed","eTestResultStatus_RESOLVED"):
            self.assertEqual(decisions([pipe(group="g"),pipe("b",160,group="g",Status=status)]),{"a":"Approved","b":"Excluded"})

    def test_incomplete_group_does_not_approve_neighbor(self):
        self.assertTrue(all(x=="Uncertain" for x in decisions([pipe(group="g"),pipe("b",160,group="g",GeomOk=False)]).values()))

    def test_fitting_not_treated_as_pipe(self):
        self.assertEqual(decisions([pipe(PairClass="перекрытие+фасонина")])["a"],"Uncertain")
        fitting=pipe("b",160,group="g",PairClass="перекрытие+фасонина",GeomOk=False)
        self.assertEqual(decisions([pipe(group="g"),fitting])["a"],"Uncertain")

    def test_malformed_and_nonfinite_data(self):
        for row in (pipe(SlabNormal=[0,0,float("nan")]),pipe(SlabNormalAngle=90),pipe(SlabFullSectionCovered=True,SlabSections=[])):
            self.assertEqual(decisions([row])["a"],"Uncertain")

    def test_polygon_cross_without_contained_vertices_has_zero_gap(self):
        self.assertEqual(polygon_gap([(-2,-.2),(2,-.2),(2,.2),(-2,.2)],[(-.2,-2),(.2,-2),(.2,2),(-.2,2)]),0)

    def test_common_plane_required(self):
        a,b=pipe(group="g"),pipe("b",160,group="g")
        b["SlabSections"][0]["origin_mm"][2]=200
        for p in b["SlabSections"][0]["contour_mm"]:p[2]=200
        self.assertTrue(all(x=="Uncertain" for x in decisions([a,b]).values()))


if __name__ == "__main__":
    unittest.main()
