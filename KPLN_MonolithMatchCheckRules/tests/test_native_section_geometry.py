import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from native_section_geometry import native_section_from_loops

def square(x,y,size,clockwise=False):
    p=np.array([[x,y],[x+size,y],[x+size,y+size],[x,y+size]],float)
    return p[::-1] if clockwise else p

class NativeContourTests(unittest.TestCase):
    def test_native_hole_and_material_island(self):
        shape,proof=native_section_from_loops([square(0,0,10),square(2,2,6,True),square(4,4,2)],68.)
        self.assertAlmostEqual(shape.area,68.);self.assertEqual(len(proof['native_holes_preserved']),1)
    def test_tiny_hole_cannot_be_filled_under_area_budget(self):
        outer=square(50000,40000,100);hole=square(50010,40010,1e-6,True)
        shape,proof=native_section_from_loops([outer,hole],10000.-1e-12)
        self.assertEqual(len(proof['native_holes_preserved']),1)
        self.assertGreater(proof['native_holes_preserved'][0]['native_void_area_mm2'],0.)
        self.assertEqual(len(shape.interiors if shape.geom_type=='Polygon' else shape.geoms[0].interiors),1)
    def test_retraced_native_contour_uses_only_original_segments(self):
        loop=np.array([[0,0],[10,0],[10,10],[0,10],[0,0],[20,0],[30,0],[30,10],[20,10],[20,0]],float)
        shape,proof=native_section_from_loops([loop],200.,True)
        self.assertAlmostEqual(shape.area,200.);self.assertEqual(len(proof['noded_contours']),1)
    def test_wrong_native_area_rejected(self):
        with self.assertRaises(ValueError):native_section_from_loops([square(0,0,10)],50.)
    def test_hole_without_material_parent_rejected(self):
        with self.assertRaises(ValueError):native_section_from_loops([square(0,0,10),square(20,20,2,True)],96.)
if __name__=='__main__':unittest.main()
