import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from surface_projection import measurement_frame,lift_triangle

class ProjectionTests(unittest.TestCase):
    def test_oblique_frame_and_source_lift(self):
        tri=np.array([[20000.,-30000.,40000.],[20007.,-29991.,40003.],[20003.,-29996.,40012.]])
        origin,basis=measurement_frame([tri])
        np.testing.assert_allclose(basis@basis.T,np.eye(3),atol=1e-14)
        bary=np.array([[.2,.3,.5],[.5,.5,0],[0,0,1.]])
        desired=bary@tri;xy=(desired-origin)@basis[:2].T
        result,error=lift_triangle(xy,tri,origin,basis)
        np.testing.assert_allclose(result,desired,atol=1e-10);self.assertLess(error,1e-8)

    def test_skinny_input_uses_stable_representative(self):
        regular=np.array([[0.,0.,0.],[10.,4.,3.],[3.,9.,5.]])
        skinny=np.array([regular[0],regular[1],regular[1]*.5+regular[2]*1e-15])
        _,basis=measurement_frame([skinny,regular])
        np.testing.assert_allclose(basis@basis.T,np.eye(3),atol=1e-14)

    def test_unstable_plane_is_explicitly_unresolved(self):
        with self.assertRaises(ValueError):measurement_frame([[[0.,0.,0.],[1.,0.,0.],[.5,1e-17,0.]]])

    def test_outside_partition_cannot_be_lifted_as_source(self):
        tri=np.array([[0.,0.,0.],[1.,0.,0.],[0.,1.,0.]])
        o,b=measurement_frame([tri]);xy=(np.array([[2.,2.,0.],[2.,3.,0.],[3.,2.,0.]])-o)@b[:2].T
        with self.assertRaises(ValueError):lift_triangle(xy,tri,o,b)

if __name__=='__main__':unittest.main()
