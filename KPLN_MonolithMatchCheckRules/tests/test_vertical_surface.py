import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from shapely.geometry import box
from vertical_surface import *

class VerticalSurfaceTests(unittest.TestCase):
    def candidate(self,z,shape=None,plane=None):
        return {'key':str(z),'shape':box(0,0,100,100) if shape is None else shape,
            'plane':np.array([.5,0.,z]) if plane is None else np.array(plane),'surface_error':0.}
    def test_equal_twenty_allowed(self):
        rows,gap,p=measure(box(0,0,100,100),np.array([.5,0,0]),[self.candidate(20)],20)
        self.assertEqual(rows[0]['over'].area,0);self.assertEqual(rows[0]['within'].area,10000)
    def test_vertical_twenty_two_not_normal_distance(self):
        rows,gap,p=measure(box(0,0,100,100),np.array([2.,0,0]),[self.candidate(22,plane=[2.,0,22])],20)
        self.assertEqual(rows[0]['over'].area,10000)
        self.assertLess(22/np.sqrt(5),15)
    def test_shifted_contour_not_height_to_distant_floor(self):
        rows,gap,p=measure(box(0,0,100,100),np.array([.5,0,0]),[self.candidate(0,box(5,0,105,100))],20)
        self.assertEqual(gap.area,500);self.assertEqual(rows[0]['over'].area,0)
    def test_opposite_run_is_not_counterpart(self):
        rows,gap,p=measure(box(0,0,100,100),np.array([.5,0,0]),[self.candidate(0,plane=[-.5,0,50])],20)
        self.assertEqual(rows,[]);self.assertEqual(gap.area,10000)
    def test_nearest_affine_surface_across_whole_triangle(self):
        source=box(0,0,100,100)
        rows,gap,p=measure(source,np.array([.5,0,0]),[self.candidate(0,plane=[.6,0,0]),self.candidate(0,plane=[.4,0,10])],20)
        self.assertAlmostEqual(sum(r['shape'].area for r in rows),10000)
        self.assertEqual(gap.area,0)
    def test_uncertainty_is_not_rounded_to_allowed(self):
        rows,gap,p=measure(box(0,0,100,100),np.array([.5,0,0]),[self.candidate(20)],20,source_error=.1)
        self.assertEqual(rows[0]['uncertain'].area,10000)

if __name__=='__main__':unittest.main()
