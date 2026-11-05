import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from parent_cover_euclidean import parent_coverage

class CoverageTests(unittest.TestCase):
    def test_normal_offset_near_full_distance_budget(self):
        p=np.array([[0.,0.,0.],[1.,0.,0.],[0.,1.,0.]])
        self.assertTrue(parent_coverage(p,[p+[0,0,.9]],1.)['passed'])
    def test_outside_euclidean_ball_is_not_covered(self):
        p=np.array([[0.,0.,0.],[1.,0.,0.],[0.,1.,0.]])
        self.assertFalse(parent_coverage(p,[p+[.7,.7,.9]],1.)['passed'])
    def test_each_child_has_its_own_normal_bound(self):
        p=np.array([[0.,0.,0.],[1.,0.,0.],[0.,1.,0.]])
        self.assertTrue(parent_coverage(p,[p+[0,0,.1],p+[0,0,10.]],1.)['passed'])
    def test_oblique_very_thin_parent_keeps_orthonormal_frame(self):
        p=np.array([[40000.,30000.,10000.],[41000.,31000.,10000.],[40500.,30500.+1e-8,10000.]])
        self.assertTrue(parent_coverage(p,[p],1e-5)['passed'])
if __name__=='__main__':unittest.main()
