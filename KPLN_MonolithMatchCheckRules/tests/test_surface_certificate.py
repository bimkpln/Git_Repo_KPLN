import unittest,tempfile,subprocess,struct,sys
from pathlib import Path
from importlib.util import spec_from_file_location,module_from_spec
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
DT=np.dtype([('parent','<i4'),('state','<i4'),('tri','<f8',(3,3)),('minimum','<f8'),('maximum','<f8'),('p','<f8',(3,)),('q','<f8',(3,))])

class NativeCertificateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp=tempfile.TemporaryDirectory();cls.folder=Path(cls.temp.name)
        spec=spec_from_file_location('build_certificate',ROOT/'scripts/build-distance-certificate.py');module=module_from_spec(spec);spec.loader.exec_module(module)
        cls.exe=module.build(cls.folder/'proof.exe')

    @classmethod
    def tearDownClass(cls):cls.temp.cleanup()

    def run_proof(self,source,tolerance,margin):
        target=np.array([[[-100,-100,0],[300,-100,0],[-100,300,0]]],float)
        for name,triangles in [('target',target),('source',source)]:
            with (self.folder/(name+'.bin')).open('wb') as f:f.write(struct.pack('<i',len(triangles)));triangles.astype('<f8').tofile(f)
        subprocess.run([str(self.exe),str(self.folder/'target.bin'),str(self.folder/'source.bin'),str(self.folder/'result.bin'),str(tolerance),'deep','native-csg-difference',str(margin)],check=True,capture_output=True)
        return np.fromfile(self.folder/'result.bin',dtype=DT,offset=4)

    def parallel_faces(self,tolerance,margin):
        tri=np.array([[[0,0,0],[10,0,0],[0,10,0]]],float);source=np.repeat(tri,3,axis=0);source[:,:,2]=np.array([tolerance-.01,tolerance,tolerance+.01])[:,None]
        return self.run_proof(source,tolerance,margin)

    def test_15mm_equality_allowed(self):self.assertEqual(self.parallel_faces(15,0)['state'].tolist(),[0,0,1])
    def test_20mm_equality_allowed(self):self.assertEqual(self.parallel_faces(20,0)['state'].tolist(),[0,0,1])
    def test_100mm_profile_is_parameter(self):self.assertEqual(self.parallel_faces(100,0)['state'].tolist(),[0,0,1])
    def test_uncertainty_never_expands_tolerance(self):self.assertEqual(self.parallel_faces(15,.02)['state'].tolist(),[2,2,2])

    def bounds(self,target,source):
        for name,triangles in [('target',target),('source',source)]:
            with (self.folder/(name+'.bin')).open('wb') as f:
                triangles=np.asarray(triangles,float);f.write(struct.pack('<i',len(triangles)));triangles.astype('<f8').tofile(f)
        subprocess.run([str(self.exe),str(self.folder/'target.bin'),str(self.folder/'source.bin'),str(self.folder/'bounds.bin'),'bounds'],check=True,capture_output=True)
        dtype=np.dtype([('minimum','<f8'),('upper','<f8'),('witness','<f8'),('face','<i4'),('p','<f8',(3,)),('q','<f8',(3,))])
        return np.fromfile(self.folder/'bounds.bin',dtype=dtype,offset=4)

    def test_whole_triangle_bounds_equal_parallel_distance(self):
        target=[[[-100,-100,0],[300,-100,0],[-100,300,0]]]
        r=self.bounds(target,[[[0,0,20],[10,0,20],[0,10,20]]])
        self.assertAlmostEqual(r['minimum'][0],20);self.assertAlmostEqual(r['upper'][0],20)

    def test_minimum_finds_interior_crossing_without_close_vertices(self):
        target=[[[-10,-10,0],[10,-10,0],[0,10,0]]]
        r=self.bounds(target,[[[-20,0,-20],[20,0,-20],[0,0,20]]])
        self.assertEqual(r['minimum'][0],0);self.assertGreater(r['upper'][0],0)

    def test_upper_does_not_use_vertexwise_nearest_as_global_maximum(self):
        source=np.array([[[0,0,0],[100,0,0],[0,100,0]]],float)
        target=np.array([[[0,0,0],[1,0,0],[0,1,0]],[[100,0,0],[99,0,0],[100,1,0]],[[0,100,0],[0,99,0],[1,100,0]]],float)
        r=self.bounds(target,source)
        self.assertGreater(r['upper'][0],40)

    def test_nearly_parallel_vertical_edge_has_no_false_intersection(self):
        source=[[[34749.97,40947.67,77130.0],[34783.00543771536,40966.866651355165,77189.67094098452],[34749.97,40947.67,77189.65307757171]]]
        target=[[[34736.71,40955.75,77205.0],[34749.97,40947.67,77205.0],[34749.97,40947.67,77280.0]]]
        r=self.bounds(target,source)
        self.assertGreater(r['minimum'][0],15.34)
        self.assertLess(r['minimum'][0],15.35)

    def test_coplanar_overlapping_triangles_still_have_zero_minimum(self):
        r=self.bounds([[[0,0,0],[10,0,0],[0,10,0]]],[[[5,-5,0],[15,5,0],[-5,5,0]]])
        self.assertLess(r['minimum'][0],1e-8)

    def test_shallow_true_intersection_is_preserved(self):
        r=self.bounds([[[-100,-100,0],[100,-100,0],[0,100,0]]],[[[-80,0,-1e-8],[80,0,1e-8],[80,0,30]]])
        self.assertLess(r['minimum'][0],1e-7)

    def test_tiny_and_zero_area_measurement_domains_are_not_silently_dropped(self):
        source=np.array([[[0,0,5],[10,0,5],[10,1e-12,5]],[[1,1,5],[1,1,5],[1,1,5]],[[0,0,5],[10,0,5],[20,0,5]]],float)
        r=self.run_proof(source,15,0)
        self.assertEqual(sorted(set(r['parent'])),[0,1,2]);self.assertTrue(np.all(r['state']==0))

    def test_crossing_face_coverage_and_bounded_work(self):
        source=np.array([[[0,0,0],[10,0,0],[0,10,30]]],float);r=self.run_proof(source,15,.02)
        area=np.linalg.norm(np.cross(r['tri'][:,1]-r['tri'][:,0],r['tri'][:,2]-r['tri'][:,0]),axis=1)/2
        self.assertAlmostEqual(float(area.sum()),float(np.linalg.norm(np.cross(source[0,1]-source[0,0],source[0,2]-source[0,0]))/2),places=7)
        self.assertLess(len(r),2048)
        self.assertTrue(np.all(r['tri'][r['state']==0,:,2]<=14.98+1e-9))
        self.assertTrue(np.all(r['tri'][r['state']==1,:,2]>=15.02-1e-9))
