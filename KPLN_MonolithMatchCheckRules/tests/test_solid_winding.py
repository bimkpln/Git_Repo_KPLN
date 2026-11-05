import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
import trimesh
from solid_winding import winding_numbers

class WindingTests(unittest.TestCase):
    def test_inside_and_outside(self):
        m=trimesh.creation.box(extents=[2,2,2]);w=winding_numbers([[0,0,0],[3,0,0]],m.vertices,m.faces)
        np.testing.assert_allclose(w,[1,0],atol=1e-12)
    def test_actual_negative_cavity_shell_remains_free(self):
        outer=trimesh.creation.box(extents=[4,4,4]);inner=trimesh.creation.box(extents=[2,2,2])
        v=np.r_[outer.vertices,inner.vertices];f=np.r_[outer.faces,inner.faces[:,::-1]+len(outer.vertices)]
        np.testing.assert_allclose(winding_numbers([[0,0,0],[1.5,0,0],[4,0,0]],v,f),[0,1,0],atol=1e-12)
    def test_translation_and_chunking(self):
        m=trimesh.creation.box(extents=[2,2,2]);shift=np.array([1e8,-2e8,3e8]);p=np.array([[0,0,0],[2,3,4],[-.5,.2,0]])
        np.testing.assert_allclose(winding_numbers(p+shift,m.vertices+shift,m.faces,point_chunk=1,triangle_chunk=3),[1,0,1],atol=1e-10)
    def test_vertex_never_gives_an_outside_decision(self):
        m=trimesh.creation.box()
        with self.assertRaises(ValueError):winding_numbers(m.vertices[:1],m.vertices,m.faces)
if __name__=='__main__':unittest.main()
