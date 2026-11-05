import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from shapely.geometry import box,Polygon
from rational_coverage import coverage

class CoverageTests(unittest.TestCase):
    def test_whole_query_in_material(self):
        self.assertTrue(coverage(box(1,1,2,2),[box(0,0,3,3)])['covered'])
    def test_material_hole_is_not_covered(self):
        p=Polygon(box(0,0,3,3).exterior,[box(1,1,2,2).exterior])
        self.assertFalse(coverage(box(1.2,1.2,1.8,1.8),[p])['covered'])
    def test_multiple_parts_cover_query(self):
        self.assertTrue(coverage(box(0,0,2,2),[box(0,0,1,2),box(1,0,2,2)])['covered'])
    def test_thin_positive_gap_stays_uncovered(self):
        p=Polygon([(50000.,40000.),(50001.,40000.),(50001.,40000.+1e-8)])
        self.assertFalse(coverage(p,[])['covered'])
    def test_partial_material_is_not_whole_coverage(self):
        self.assertFalse(coverage(box(0,0,2,2),[box(0,0,1,2)])['covered'])
    def test_exact_gap_smaller_than_float_spacing_is_not_welded(self):
        from fractions import Fraction as F
        e=F(1,10**30)
        left=[(F(0),F(0)),(1-e,F(0)),(1-e,F(2)),(F(0),F(2))]
        right=[(1+e,F(0)),(F(2),F(0)),(F(2),F(2)),(1+e,F(2))]
        p=coverage(box(0,0,2,2),[],material_cycles=[left,right])
        self.assertFalse(p['covered']);self.assertTrue(p['uncovered_exact_witnesses'])
    def test_projected_overlap_multiplicity_is_explicit(self):
        p=coverage(box(0,0,2,2),[box(0,0,1.5,2),box(.5,0,2,2)])
        self.assertTrue(p['covered']);self.assertEqual(p['maximum_material_winding_in_query'],2)
if __name__=='__main__':unittest.main()
