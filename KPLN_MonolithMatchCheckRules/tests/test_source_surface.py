import unittest,copy,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import numpy as np,trimesh
from source_surface import *

def indexed(mesh):
    v,ii=np.unique(mesh.triangles.reshape(-1,3),axis=0,return_inverse=True)
    return v,ii.reshape(-1,3)

class CorrespondenceTests(unittest.TestCase):
    def test_native_volume_and_three_axis_live_csg(self):
        v,f=indexed(trimesh.creation.box(extents=[30,40,50]))
        v+=np.array([68000,37000,82000])
        vv,ff,n,c=recover(v,f)
        proof=verify_native(vv,ff,n)
        self.assertTrue(proof['passed']);self.assertAlmostEqual(n.volume(),60000)
        self.assertEqual(c['maximum_surface_correspondence_bound_mm'],0)

    def test_missing_surface_never_filled(self):
        v,f=indexed(trimesh.creation.box(extents=[100,100,100]))
        with self.assertRaises(TopologyError):recover(v,f[:-2])

    def test_duplicate_internal_interfaces_preserve_signed_material(self):
        a=trimesh.creation.box(extents=[20,20,20]);b=a.copy();b.apply_translation([20,0,0])
        v,f=indexed(trimesh.util.concatenate([a,b]))
        vv,ff,n,c=recover(v,f)
        self.assertTrue(verify_native(vv,ff,n)['passed'])
        self.assertAlmostEqual(n.volume(),16000)

    def test_opposite_planes_only_cancel_overlap(self):
        v=np.array([[0,0,0],[10,0,0],[0,10,0],[10,10,0]],float)
        f=np.array([[0,1,2],[2,1,0],[1,3,2]])
        vv,ff,c=reduce_opposite_planes(v,f)
        self.assertAlmostEqual(trimesh.Trimesh(vertices=vv,faces=ff,process=False).area,50)
        self.assertEqual(c['new_caps'],0)

    def test_transitive_vertex_cluster_cannot_exceed_bound(self):
        v=np.array([[0,0,0],[.09,0,0],[.18,0,0],[0,1,0]],float)
        f=np.array([[0,1,3],[1,2,3]])
        with self.assertRaisesRegex(TopologyError,'Transitive'):sew_vertices(v,f,.1)

    def test_independent_validator_rejects_forged_reference(self):
        v=np.array([[0,0,0],[1,0,0],[0,1,0],[.5,0,0]],float);f=np.array([[0,1,2],[1,3,2]])
        vv,ff,c=split_t_seams(v,f,.001,reference_evidence=True)
        c=copy.deepcopy(c);c['reference_triangle_coordinates'][0][0]=[100,100,100]
        with self.assertRaises(TopologyError):verify_subdivision(v,f,vv,ff,c)

    def test_nonfinite_coordinates_rejected(self):
        v,f=indexed(trimesh.creation.box());v[0,0]=np.nan
        with self.assertRaisesRegex(TopologyError,'Nonfinite'):recover(v,f)

if __name__=='__main__':unittest.main()
