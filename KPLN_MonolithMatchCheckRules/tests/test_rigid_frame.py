import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
import manifold3d as md
from rigid_frame import rotation_matrix,independent_volume
class RigidFrameTests(unittest.TestCase):
    def test_frame_matches_native_rotation_and_restores_coordinates(self):
        cube=md.Manifold.cube([10,20,30],center=True);p=np.array(cube.to_mesh64().vert_properties)[:,:3];angles=[13.7,7.1,3.9];R=rotation_matrix(angles)
        q=np.array(cube.rotate(angles).to_mesh64().vert_properties)[:,:3]
        self.assertLess(np.max(abs(q-p@R.T)),1e-10);self.assertLess(np.max(abs(q@R-p)),1e-10)
    def test_operands_and_clip_rotated_together_preserve_physical_domain(self):
        source=md.Manifold.cube([10,20,30],center=True).translate([3,4,5]);clip=md.Manifold.cube([12,12,12],center=True);R=rotation_matrix([31,17,11])
        a=source^clip;b=source.transform(np.c_[R,np.zeros(3)])^clip.transform(np.c_[R,np.zeros(3)])
        self.assertAlmostEqual(a.volume(),b.volume(),places=7)
        b=b.transform(np.c_[R.T,np.zeros(3)]);self.assertLess((a-b).volume()+(b-a).volume(),1e-6)
    def test_centered_signed_volume_is_independent_of_native_volume_cache(self):
        box=md.Manifold.cube([10,20,30],center=True).translate([40000,50000,70000]);self.assertAlmostEqual(independent_volume(box),6000)
if __name__=='__main__':unittest.main()
