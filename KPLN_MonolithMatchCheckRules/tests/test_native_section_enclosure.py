import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from unittest.mock import patch
import manifold3d as md
from native_section_union import assemble_sections

class SectionEnclosureTests(unittest.TestCase):
    def test_positive_omission_rejected_without_explicit_bound(self):
        part=md.CrossSection.square([10.,10.]);missing=md.CrossSection.square([10.-5e-8,10.])
        with patch.object(md.CrossSection,'batch_boolean',return_value=missing):
            with self.assertRaisesRegex(ValueError,'omits'):assemble_sections([part])

    def test_roundoff_strip_requires_complete_distance_enclosure(self):
        part=md.CrossSection.square([10.,10.]);missing=md.CrossSection.square([10.-5e-8,10.])
        with patch.object(md.CrossSection,'batch_boolean',return_value=missing):
            _,shape,proof=assemble_sections([part],representation_bound_mm=1e-7)
        self.assertFalse(proof['exact_original_material_inclusion'])
        self.assertTrue(proof['absence_decisions_require_original_seed_and_boundary_band_audit'])
        self.assertLess(shape.area,100.)

    def test_finite_missing_strip_never_excused_by_area_budget(self):
        part=md.CrossSection.square([.001,.001]);missing=md.CrossSection.square([.00099,.001])
        with patch.object(md.CrossSection,'batch_boolean',return_value=missing):
            with self.assertRaisesRegex(ValueError,'exceeds'):assemble_sections([part],representation_bound_mm=1e-7)

    def test_extra_material_beyond_enclosure_rejected(self):
        part=md.CrossSection.square([.001,.001]);extra=md.CrossSection.square([.00101,.001])
        with patch.object(md.CrossSection,'batch_boolean',return_value=extra):
            with self.assertRaisesRegex(ValueError,'creates'):assemble_sections([part],representation_bound_mm=1e-7)

if __name__=='__main__':unittest.main()
