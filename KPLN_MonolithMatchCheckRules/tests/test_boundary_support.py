import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from boundary_support import complete_segment_support,boundary_patch_receipt

class DirectedBoundaryTests(unittest.TestCase):
    def test_full_opposite_edge_subdivision(self):
        targets=np.array([[[5,0,0],[0,0,0]],[[10,0,0],[5,0,0]]],float)
        self.assertTrue(complete_segment_support(np.array([0,0,0]),np.array([10,0,0]),targets,1e-8))
    def test_endpoints_do_not_hide_middle_gap(self):
        targets=np.array([[[2,0,0],[0,0,0]],[[10,0,0],[8,0,0]]],float)
        self.assertFalse(complete_segment_support(np.array([0,0,0]),np.array([10,0,0]),targets,1e-8))
    def test_wrong_direction_rejected(self):
        self.assertFalse(complete_segment_support(np.array([0,0,0]),np.array([10,0,0]),np.array([[[0,0,0],[10,0,0]]],float),1e-8))
    def test_closed_finish_solid_not_a_boundary_patch(self):
        import trimesh
        solid=trimesh.creation.box(extents=[1,1,1])
        proof=boundary_patch_receipt(solid.triangles,solid.triangles,1e-8)
        self.assertFalse(proof['passed']);self.assertEqual(proof['closed_patch_components'],1)

if __name__=='__main__':unittest.main()
