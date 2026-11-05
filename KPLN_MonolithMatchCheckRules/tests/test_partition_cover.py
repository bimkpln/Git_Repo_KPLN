import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from partition_cover import parent_coverage
class PartitionCoverTests(unittest.TestCase):
    def test_identical_oblique_sliver_is_covered_without_unstable_normal(self):
        import numpy as np
        base=np.array([45000.,55000.,75000.]);u=np.array([300.,400.,866.0254037844386])
        p=np.array([base,base+u,base+u*.5+np.array([-3.2e-9,2.4e-9,0.])])
        self.assertTrue(parent_coverage(p,[p],1e-5)['passed'])
    def test_exact_two_triangle_partition(self):
        p=[[0,0,0],[10,0,0],[0,10,0]];c=[[[0,0,0],[5,0,0],[0,10,0]],[[5,0,0],[10,0,0],[0,10,0]]]
        self.assertTrue(parent_coverage(p,c,1e-5)['passed'])
    def test_missing_finite_parent_region_rejected(self):
        p=[[0,0,0],[10,0,0],[0,10,0]];c=[[[0,0,0],[1,0,0],[0,1,0]]]
        self.assertFalse(parent_coverage(p,c,1e-5)['passed'])
    def test_off_plane_child_rejected(self):
        p=[[0,0,0],[10,0,0],[0,10,0]];c=[[[0,0,.001],[10,0,.001],[0,10,.001]]]
        self.assertFalse(parent_coverage(p,c,1e-5)['passed'])
    def test_degenerate_line_kept_as_measurement_domain(self):
        p=[[0,0,0],[10,0,0],[5,0,0]];c=[[[0,0,0],[5,0,0],[3,0,0]],[[5,0,0],[10,0,0],[7,0,0]]]
        self.assertTrue(parent_coverage(p,c,1e-5)['passed'])
if __name__=='__main__':unittest.main()
