import unittest,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import numpy as np,trimesh
from planar_skin import canonical_skin
from source_surface import TopologyError

class CanonicalSkinTests(unittest.TestCase):
    def test_closed_box_source_skin_preserved(self):
        m=trimesh.creation.box(extents=[20,30,40]);v,f,n,c=canonical_skin(m.vertices,m.faces)
        self.assertAlmostEqual(n.volume(),24000)
        self.assertEqual(c['new_or_guessed_caps'],0)
        self.assertTrue(c['native_acceptance']['passed'])
        self.assertLess(c['maximum_additional_surface_correspondence_bound_mm'],1e-5)

    def test_large_translation_keeps_volume(self):
        m=trimesh.creation.box(extents=[20,30,40]);m.apply_translation([68000,36000,82000])
        v,f,n,c=canonical_skin(m.vertices,m.faces)
        self.assertAlmostEqual(n.volume(),24000)

    def test_missing_source_face_cannot_be_capped(self):
        m=trimesh.creation.box(extents=[20,30,40])
        with self.assertRaises(TopologyError):canonical_skin(m.vertices,m.faces[:-2])

    def test_rigid_rotated_box_preserves_signed_coverage(self):
        m=trimesh.creation.box(extents=[20,30,40]);m.apply_transform(trimesh.transformations.rotation_matrix(.564,[0,0,1]))
        v,f,n,c=canonical_skin(m.vertices,m.faces)
        self.assertLess(abs(n.volume()-24000),3*m.area*c['maximum_additional_surface_correspondence_bound_mm'])
        self.assertLess(max(p['signed_area_error_mm2'] for p in c['planes']),1e-5)

    def test_disconnected_material_parts_preserved(self):
        a=trimesh.creation.box(extents=[10,10,10]);b=a.copy();b.apply_translation([20,0,0]);m=trimesh.util.concatenate([a,b])
        v,f,n,c=canonical_skin(m.vertices,m.faces)
        self.assertAlmostEqual(n.volume(),2000)
        self.assertEqual(len(n.decompose()),2)

    def test_degenerate_triangle_rejected(self):
        with self.assertRaises(TopologyError):canonical_skin(np.array([[0,0,0],[1,0,0],[2,0,0]],float),np.array([[0,1,2]]))

if __name__=='__main__':unittest.main()
