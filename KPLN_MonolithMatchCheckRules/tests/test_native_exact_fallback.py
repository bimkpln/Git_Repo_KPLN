import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from unittest.mock import patch
import numpy as np
from native_section_geometry import native_section_from_loops

def square(x,y,size,cw=False):
    p=np.array([[x,y],[x+size,y],[x+size,y+size],[x,y+size]],float)
    return p[::-1] if cw else p

class ExactFallbackTests(unittest.TestCase):
    def test_fallback_retains_hole_and_island(self):
        with patch('native_section_geometry.nested_section',side_effect=ValueError('nesting failure')):
            shape,proof=native_section_from_loops([square(0,0,10),square(2,2,6,True),square(4,4,2)],68.,True)
        self.assertAlmostEqual(shape.area,68.)
        self.assertFalse(shape.contains(__import__('shapely').Point(3,3)))
        self.assertTrue(shape.contains(__import__('shapely').Point(5,5)))
        self.assertEqual(proof['arrangement']['free_faces'],1)
    def test_fallback_keeps_tiny_void_at_large_coordinates(self):
        with patch('native_section_geometry.nested_section',side_effect=ValueError('nesting failure')):
            shape,_=native_section_from_loops([square(50000,40000,100),square(50010,40010,1e-6,True)],10000.-1e-12,True)
        parts=[shape] if shape.geom_type=='Polygon' else list(shape.geoms)
        self.assertEqual(sum(len(p.interiors) for p in parts),1)
    def test_fallback_rejects_negative_region_without_parent(self):
        with patch('native_section_geometry.nested_section',side_effect=ValueError('nesting failure')):
            with self.assertRaisesRegex(ValueError,'Negative contour'):
                native_section_from_loops([square(0,0,10),square(20,20,1e-6,True)],100.,True)
    def test_fallback_rejects_wrong_area(self):
        with patch('native_section_geometry.nested_section',side_effect=ValueError('nesting failure')):
            with self.assertRaisesRegex(ValueError,'area'):
                native_section_from_loops([square(0,0,10)],50.,True)
    def test_disabled_fallback_preserves_rejection(self):
        with patch('native_section_geometry.nested_section',side_effect=ValueError('nesting failure')):
            with self.assertRaisesRegex(ValueError,'nesting failure'):
                native_section_from_loops([square(0,0,10)],100.,False)

if __name__=='__main__':unittest.main()
