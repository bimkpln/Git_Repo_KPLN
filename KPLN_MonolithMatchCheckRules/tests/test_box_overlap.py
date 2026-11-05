import sys,unittest
from pathlib import Path
from fractions import Fraction as F
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from box_overlap import exact_box_overlap,domain

I=[[1,0,0],[0,1,0],[0,0,1]]
class BoxOverlapTests(unittest.TestCase):
    def test_identical(self):self.assertTrue(exact_box_overlap(I,[[0]*3,[1]*3],I,[[0]*3,[1]*3])['positive_volume_overlap'])
    def test_contained(self):self.assertTrue(exact_box_overlap(I,[[-2]*3,[2]*3],I,[[-1]*3,[1]*3])['positive_volume_overlap'])
    def test_disjoint(self):self.assertFalse(exact_box_overlap(I,[[0]*3,[1]*3],I,[[2,0,0],[3,1,1]])['positive_volume_overlap'])
    def test_face_touch_is_not_positive_volume(self):self.assertFalse(exact_box_overlap(I,[[0]*3,[1]*3],I,[[1,0,0],[2,1,1]])['positive_volume_overlap'])
    def test_edge_touch(self):self.assertFalse(exact_box_overlap(I,[[0]*3,[1]*3],I,[[1,1,0],[2,2,1]])['positive_volume_overlap'])
    def test_sub_float_spacing_overlap_retained(self):
        e=F(1,2**80);self.assertTrue(exact_box_overlap(I,[[0]*3,[1]*3],I,[[1-e,0,0],[2,1,1]])['positive_volume_overlap'])
    def test_oblique_basis_is_not_orthogonalized(self):
        skew=[[1,2,0],[0,1,0],[0,0,1]];d=domain(skew,[[0]*3,[1]*3])
        self.assertIn((F(-2),F(1),F(0)),d['vertices'])
        self.assertTrue(exact_box_overlap(skew,[[0]*3,[1]*3],I,[[-1,F(1,4),0],[0,F(3,4),1]])['positive_volume_overlap'])
    def test_rotated(self):
        R=[[F(3,5),F(4,5),0],[-F(4,5),F(3,5),0],[0,0,1]]
        self.assertTrue(exact_box_overlap(R,[[-1]*3,[1]*3],I,[[-1]*3,[1]*3])['positive_volume_overlap'])
    def test_bad_domains_rejected(self):
        for B,bb in [([[0]*3]*3,[[0]*3,[1]*3]),(I,[[0]*3,[1,1,0]])]:
            with self.assertRaises(ValueError):exact_box_overlap(B,bb,I,[[0]*3,[1]*3])
    def test_symmetry(self):
        R=[[1,1,0],[0,1,0],[0,0,1]];a=[[-1]*3,[1]*3];b=[[2,2,0],[3,3,1]]
        self.assertEqual(exact_box_overlap(I,a,R,b)['positive_volume_overlap'],exact_box_overlap(R,b,I,a)['positive_volume_overlap'])
if __name__=='__main__':unittest.main()
