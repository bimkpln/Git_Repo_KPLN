import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from unittest.mock import patch
import manifold3d as md
import native_section_union as target

class IndependentUnionTests(unittest.TestCase):
    def test_rejected_native_contours_use_valid_original_parts(self):
        left=md.CrossSection.square([1.,2.]);right=left.translate([1.,0.])
        original=target.section_from_loops;counter=[0]
        def reject_first(*a,**k):
            counter[0]+=1
            if counter[0]==1:raise ValueError('unrepresentable native contour')
            return original(*a,**k)
        with patch.object(target,'section_from_loops',side_effect=reject_first):
            _,shape,p=target.assemble_sections([left,right])
        self.assertAlmostEqual(shape.area,4.)
        self.assertEqual(p['native_reconstruction']['method'],'independent-valid-source-material-union')
    def test_independent_parts_cannot_override_wrong_native_area(self):
        original=target.section_from_loops;counter=[0]
        def reject_first(*a,**k):
            counter[0]+=1
            if counter[0]==1:raise ValueError('unrepresentable native contour')
            return original(*a,**k)
        with patch.object(target,'section_from_loops',side_effect=reject_first):
            with patch.object(md.CrossSection,'batch_boolean',return_value=md.CrossSection.square([3.,3.])):
                with self.assertRaisesRegex(ValueError,'native area'):
                    target.assemble_sections([md.CrossSection.square([1.,1.])])
if __name__=='__main__':unittest.main()
