import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
import manifold3d as md
from box_material_certificate import box_material_certificate
class BoxMaterialTests(unittest.TestCase):
    def test_enclosing_body_proves_occupied_instead_of_geometry_failure(self):
        domain=md.Manifold.cube([2,2,2],center=True);body=md.Manifold.cube([10,10,10],center=True)
        self.assertAlmostEqual((domain^body).volume(),domain.volume())
        p=box_material_certificate(np.ones(12),False,'inside',.1,.01,True);self.assertEqual(p['state'],'occupied')
    def test_external_body_proves_empty(self):
        p=box_material_certificate(np.ones(12),False,'outside',.1,.01,True);self.assertEqual(p['state'],'empty')
    def test_internal_component_keeps_membership_nonconstant(self):
        p=box_material_certificate(np.ones(12),True,'outside',.1,.01,True);self.assertEqual(p['state'],'unproved')
    def test_unknown_membership_never_becomes_empty(self):
        p=box_material_certificate(np.ones(12),False,'uncertain',.1,.01,True);self.assertEqual(p['state'],'unproved')
    def test_contact_does_not_prove_constant_state(self):
        p=box_material_certificate(np.zeros(12),False,'inside',.1,.01,True);self.assertEqual(p['state'],'unproved')
if __name__=='__main__':unittest.main()
