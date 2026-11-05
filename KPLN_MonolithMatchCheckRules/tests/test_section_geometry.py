import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from shapely.geometry import box,Polygon
from section_geometry import section_from_loops

class SourceContourTests(unittest.TestCase):
    def test_nested_islands_and_holes(self):
        rings=[box(0,0,10,10).exterior.coords,box(1,1,9,9).exterior.coords,box(2,2,8,8).exterior.coords,box(3,3,7,7).exterior.coords]
        shape,proof=section_from_loops(rings,expected_area=100-64+36-16)
        self.assertEqual(shape.area,56)
        self.assertEqual(proof['new_or_guessed_contours'],0)

    def test_disjoint_material_is_preserved(self):
        shape,proof=section_from_loops([box(0,0,2,2).exterior.coords,box(3,0,5,2).exterior.coords],8)
        self.assertEqual(shape.area,8)

    def test_incorrect_native_area_is_rejected(self):
        with self.assertRaises(ValueError):section_from_loops([box(0,0,2,2).exterior.coords],5)

    def test_invalid_contour_not_filled(self):
        with self.assertRaises(ValueError):section_from_loops([[(0,0),(3,3),(0,3),(2,0)]])

    def test_many_holes_match_exact_area(self):
        holes=[box(x+.2,y+.2,x+.8,y+.8).exterior.coords for x in range(20) for y in range(20)]
        shape,proof=section_from_loops([box(0,0,20,20).exterior.coords]+holes,400-400*.36)
        self.assertAlmostEqual(shape.area,256)

    def test_native_touching_lobes_use_only_existing_edges(self):
        loop=[(0,0),(1,0),(1,1),(0,1),(0,0),(-1,0),(-1,-1),(0,-1),(0,0)]
        shape,proof=section_from_loops([loop],2,allow_noded_contours=True)
        self.assertEqual(shape.area,2)
        self.assertEqual(proof['noded_touching_contours'],[0])
        self.assertEqual(proof['new_or_guessed_contours'],0)

if __name__=='__main__':unittest.main()
