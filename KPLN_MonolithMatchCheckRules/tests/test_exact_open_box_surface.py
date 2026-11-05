import unittest
from fractions import Fraction as F
import numpy as np
from exact_open_box_surface import triangle_intersects_open_box as hit,source_surface_intersects_open_box as surface
L=[0,0,0];H=[1,1,1]
class ExactOpenBoxTests(unittest.TestCase):
    def test_face_contact_is_not_interior(self):self.assertFalse(hit([[0,0,0],[0,1,0],[0,0,1]],L,H))
    def test_crossing_triangle_with_all_input_vertices_outside(self):self.assertTrue(hit([[-2,.5,.5],[2,-2,.5],[2,2,.5]],L,H))
    def test_edge_touch_only(self):self.assertFalse(hit([[0,0,0],[-1,0,1],[0,-1,1]],L,H))
    def test_sub_float_spacing_positive_interior_is_retained(self):
        e=F(1,2**100);self.assertTrue(hit([[e,0,0],[e,1,0],[e,0,1]],L,H))
    def test_triangle_beyond_corner_with_overlapping_AABB(self):self.assertFalse(hit([[2,.9,.9],[.9,2,.9],[.9,.9,2]],L,H))
    def test_triangle_inside_box(self):self.assertTrue(hit([[.2,.2,.2],[.8,.2,.2],[.2,.8,.2]],L,H))
    def test_degenerate_query_rejected(self):
        with self.assertRaises(ValueError):hit([[0,0,0]]*3,L,[1,1,0])
    def test_exact_transform_preserves_boundary(self):
        v=np.array([[0,0,0],[0,1,0],[0,0,1.]])
        self.assertFalse(surface(v,[[0,1,2]],np.eye(3),L,H)['intersects'])
    def test_oblique_transform_is_literal(self):
        v=np.array([[0,.2,.2],[0,.8,.2],[0,.2,.8]])
        self.assertTrue(surface(v,[[0,1,2]],[[1,1,0],[0,1,0],[0,0,1]],L,H)['intersects'])
if __name__=='__main__':unittest.main()
