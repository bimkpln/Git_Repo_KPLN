import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
import manifold3d as md
from box_empty_certificate import empty_box_certificate

class EmptyBoxTests(unittest.TestCase):
    def setUp(self):self.domain=md.Manifold.cube([10,10,10],center=True)
    def certify(self,source,lower,inside,outside,covered=True):
        actual=(self.domain^source).volume()
        certificate=empty_box_certificate(np.full(12,lower),inside,outside,.1,.01,covered)
        return actual,certificate
    def test_separated_material_has_empty_intersection(self):
        actual,p=self.certify(md.Manifold.cube([1,1,1],center=True).translate([20,0,0]),1.,False,True)
        self.assertEqual(actual,0);self.assertTrue(p['passed'])
    def test_nested_off_center_component_cannot_pass_boundary_distance_only(self):
        actual,p=self.certify(md.Manifold.cube([1,1,1],center=True).translate([2,0,0]),1.,True,True)
        self.assertGreater(actual,0);self.assertFalse(p['passed'])
    def test_box_inside_material_cannot_pass_vertex_exclusion_only(self):
        actual,p=self.certify(md.Manifold.cube([20,20,20],center=True),1.,False,False)
        self.assertGreater(actual,0);self.assertFalse(p['passed'])
    def test_crossing_bar_without_internal_vertices_or_center_membership(self):
        actual,p=self.certify(md.Manifold.cube([1,20,1],center=True).translate([2,0,0]),0.,False,True)
        self.assertGreater(actual,0);self.assertFalse(p['passed'])
    def test_equal_clearance_is_not_a_strict_separation_certificate(self):
        p=empty_box_certificate(np.full(12,.1+.01),False,True,.1,.01,True)
        self.assertFalse(p['passed'])
    def test_incomplete_source_coverage_rejected(self):
        p=empty_box_certificate(np.ones(12),False,True,.1,.01,False)
        self.assertFalse(p['passed'])
    def test_missing_boundary_faces_rejected(self):
        with self.assertRaises(ValueError):empty_box_certificate(np.ones(11),False,True,.1,.01,True)
if __name__=='__main__':unittest.main()
