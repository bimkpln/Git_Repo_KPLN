import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from unittest.mock import patch
import manifold3d as md
from native_section_union import assemble_sections

class InputVoidTests(unittest.TestCase):
    def test_complete_small_hole_is_never_excused_by_representation_budget(self):
        outer=md.CrossSection.square([10.,10.]);hole=md.CrossSection.square([1e-6,1e-6]).translate([5.,5.]);part=outer-hole
        with patch.object(md.CrossSection,'batch_boolean',return_value=outer):
            with self.assertRaisesRegex(ValueError,'void'):assemble_sections([part],representation_bound_mm=1e-5)
    def test_real_other_material_can_fill_source_host_hole(self):
        outer=md.CrossSection.square([10.,10.]);hole=md.CrossSection.square([1.,1.]).translate([5.,5.])
        native,shape,p=assemble_sections([outer-hole,hole])
        self.assertAlmostEqual(shape.area,100.)
if __name__=='__main__':unittest.main()
