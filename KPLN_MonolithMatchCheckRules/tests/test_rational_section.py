import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
import trimesh
from rational_section import section_cycles

class RationalSectionTests(unittest.TestCase):
    def test_box_has_one_positive_exact_cycle(self):
        m=trimesh.creation.box(extents=[2,2,2]);loops,p=section_cycles(m.vertices,m.faces,np.eye(3),0.)
        self.assertEqual(len(loops),1);self.assertEqual(loops[0]['signed_area'],4)
    def test_reflection_preserves_material_orientation(self):
        m=trimesh.creation.box(extents=[2,2,2]);loops,_=section_cycles(m.vertices,m.faces,np.diag([-1,1,1]),0.)
        self.assertEqual(loops[0]['signed_area'],4)
    def test_through_hole_has_negative_exact_cycle(self):
        m=trimesh.creation.annulus(r_min=1,r_max=2,height=2,sections=4);loops,_=section_cycles(m.vertices,m.faces,np.eye(3),0.)
        self.assertEqual(sum(q['signed_area']<0 for q in loops),1)
        self.assertAlmostEqual(float(sum(q['signed_area'] for q in loops)),6.,places=12)
    def test_plane_through_vertex_is_not_silently_shifted(self):
        m=trimesh.creation.box(extents=[2,2,2])
        with self.assertRaisesRegex(ValueError,'one-sided'):section_cycles(m.vertices,m.faces,np.eye(3),1.)
    def test_explicit_positive_side_limit_at_box_boundaries(self):
        m=trimesh.creation.box(extents=[2,2,2])
        bottom,p=section_cycles(m.vertices,m.faces,np.eye(3),-1.,positive_side_limit=True)
        top,q=section_cycles(m.vertices,m.faces,np.eye(3),1.,positive_side_limit=True)
        self.assertEqual(sum(r['signed_area'] for r in bottom),4)
        self.assertEqual(sum(r['signed_area'] for r in top),0)
        self.assertTrue(p['positive_side_limit_applied']);self.assertTrue(q['positive_side_limit_applied'])
if __name__=='__main__':unittest.main()
