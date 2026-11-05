import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import manifold3d as md
from shapely.geometry import Polygon
from native_section_union import *
class OriginalSectionUnionTests(unittest.TestCase):
    def rectangle(self,x,y,w,h):return md.CrossSection.square([w,h]).translate([x,y])
    def test_overlapping_material_unioned(self):
        n,s,p=assemble_sections([self.rectangle(0,0,10,10),self.rectangle(5,0,10,10)])
        self.assertEqual(s.area,150);self.assertTrue(p['passed'])
    def test_submillimetre_hole_is_preserved(self):
        outer=self.rectangle(0,0,10,10);hole=self.rectangle(4.999,4.999,.002,.002)
        n,s,p=assemble_sections([outer-hole]);self.assertEqual(len(s.interiors),1)
        self.assertGreater(Polygon(s.interiors[0]).area,0);self.assertFalse(p['hole_size_filter_applied'])
    def test_material_island_not_misreported_as_free_space(self):
        frame=self.rectangle(0,0,10,10)-self.rectangle(1,1,8,8);island=self.rectangle(4,4,2,2)
        n,s,p=assemble_sections([frame,island]);self.assertEqual(s.area,40)
    def test_corner_touching_material_keeps_both_lobes(self):
        n,s,p=assemble_sections([self.rectangle(0,0,2,2),self.rectangle(2,2,2,2)]);self.assertEqual(s.area,8)
    def test_empty_source_has_no_invented_section(self):
        n,s,p=assemble_sections([]);self.assertTrue(s.is_empty);self.assertEqual(n.area(),0)
if __name__=='__main__':
    unittest.main()
