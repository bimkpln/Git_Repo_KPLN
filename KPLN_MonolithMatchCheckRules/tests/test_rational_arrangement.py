import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from fractions import Fraction as F
from rational_arrangement import positive_region

def square(x,y,size,negative=False):
    p=[(F(x),F(y)),(F(x)+F(size),F(y)),(F(x)+F(size),F(y)+F(size)),(F(x),F(y)+F(size))]
    return p[::-1] if negative else p

class ArrangementTests(unittest.TestCase):
    def test_overlapping_positive_contours_are_material_union(self):
        shape,p=positive_region([square(0,0,10),square(5,5,10)])
        self.assertEqual(shape.area,175.);self.assertIn(2,p['face_windings'])
    def test_native_hole_and_island(self):
        shape,p=positive_region([square(0,0,10),square(2,2,6,True),square(4,4,2)])
        self.assertEqual(shape.area,68.);self.assertGreater(p['free_faces'],0)
    def test_tiny_hole_is_retained(self):
        shape,p=positive_region([square(0,0,10),square(2,2,F(1,10**8),True)])
        self.assertEqual(len(shape.interiors),1);self.assertGreater(p['free_faces'],0)
    def test_opposite_internal_edges_cancel_without_caps(self):
        shape,p=positive_region([square(0,0,1),square(1,0,1)])
        self.assertEqual(shape.area,2.);self.assertEqual(p['graph_components'],1)
if __name__=='__main__':unittest.main()
