import unittest,tempfile,struct,hashlib,json
from pathlib import Path
from fractions import Fraction as F
from nwd_local_coordinates import validate_local_reference,exact_affine_point,callback_point

IDENTITY=[1.,0,0,0,0,1.,0,0,0,0,1.,0,0,0,0,1.]

class LocalSourceTests(unittest.TestCase):
    def fixture(self,p,points=None,transform=None):
        points=points or [(0.,0.,0.),(1.,0.,0.),(0.,1.,0.)];m=transform or IDENTITY.copy()
        local=b''.join(struct.pack('<3d',*x) for x in points);world=b''.join(struct.pack('<3d',*callback_point(x,m)) for x in points)
        (p/'nwd_local_triangles.bin').write_bytes(local);(p/'nwd_triangles.bin').write_bytes(world)
        return {'local_triangle_file':'nwd_local_triangles.bin','local_triangle_sha256':hashlib.sha256(local).hexdigest()},[
            {'triangle_offset':0,'triangle_count':1,'local_to_world':m}]
    def test_legacy_without_sidecar_is_supported(self):
        with tempfile.TemporaryDirectory() as d:self.assertIsNone(validate_local_reference(d,{},[]))
    def test_complete_literal_source_recomposition(self):
        with tempfile.TemporaryDirectory() as d:
            m,f=self.fixture(Path(d));r=validate_local_reference(d,m,f)
            self.assertEqual(r['vertices'],3);self.assertTrue(r['passed']);self.assertFalse(r['release_allowed'])
    def test_far_translation_retains_source_point_lost_in_world_float(self):
        m=IDENTITY.copy();m[12]=float(2**54)
        self.assertEqual(callback_point((1.,0.,0.),m)[0],float(2**54))
        self.assertEqual(exact_affine_point((1.,0.,0.),m)[0],F(2**54+1))
        with tempfile.TemporaryDirectory() as d:
            manifest,fragments=self.fixture(Path(d),transform=m)
            self.assertTrue(validate_local_reference(d,manifest,fragments)['passed'])
    def test_affine_matrix_orientation_and_translation_are_literal(self):
        m=[0.,1.,0.,0.,-1.,0.,0.,0.,0.,0.,1.,0.,10.,20.,30.,1.]
        self.assertEqual(exact_affine_point((2.,3.,4.),m),(F(7),F(22),F(34)))
    def test_hash_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            m,f=self.fixture(Path(d));m['local_triangle_sha256']='0'*64
            with self.assertRaisesRegex(ValueError,'sidecar'):validate_local_reference(d,m,f)
    def test_truncated_local_coordinates_are_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);m,f=self.fixture(p);(p/'nwd_local_triangles.bin').write_bytes(b'')
            with self.assertRaisesRegex(ValueError,'sidecar'):validate_local_reference(d,m,f)
    def test_wrong_matrix_cannot_double_transform_world_coordinates(self):
        with tempfile.TemporaryDirectory() as d:
            m,f=self.fixture(Path(d));f[0]['local_to_world'][12]=1.
            with self.assertRaisesRegex(ValueError,'callback'):validate_local_reference(d,m,f)
    def test_unaccounted_fragment_bytes_are_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            m,f=self.fixture(Path(d));f[0]['triangle_count']=0
            with self.assertRaisesRegex(ValueError,'unaccounted'):validate_local_reference(d,m,f)
    def test_nonfinite_local_source_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);m,f=self.fixture(p);data=struct.pack('<9d',float('nan'),0,0,1,0,0,0,1,0)
            (p/'nwd_local_triangles.bin').write_bytes(data);m['local_triangle_sha256']=hashlib.sha256(data).hexdigest()
            with self.assertRaisesRegex(ValueError,'Non-finite'):validate_local_reference(d,m,f)
    def test_sidecar_path_cannot_escape_export_directory(self):
        with tempfile.TemporaryDirectory() as d:
            m,f=self.fixture(Path(d));m['local_triangle_file']='../source.bin'
            with self.assertRaisesRegex(ValueError,'file name'):validate_local_reference(d,m,f)
    def test_undeclared_sidecar_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);m,f=self.fixture(p)
            with self.assertRaisesRegex(ValueError,'SHA256'):validate_local_reference(d,{},f)
    def test_projective_matrix_is_rejected(self):
        m=IDENTITY.copy();m[3]=1e-13
        with self.assertRaisesRegex(ValueError,'affine'):exact_affine_point((0,0,0),m)
    def test_import_receipt_binds_optional_source_coordinates(self):
        from nwd_import import validate_export
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);manifest,fragments=self.fixture(p)
            manifest.update(complete=True,schema=2,input_sha256='a'*64,geometry_nodes=1,scene_nodes=1)
            node={'index':0,'offset':0,'triangles':1,'fragments':1,'excludedBySet':False,
                'ancestry':[{'name':'part','properties':[{'category':'Объект','name':'Id','value':'123'}]},
                            {'name':'source.nwc','properties':[]}],
                'fragmentRecords':[{'fragment_id':'0:0','offset':0,'triangles':1,
                    'local_to_world':IDENTITY.copy(),'path_matches_node':True,'excludedBySet':False,
                    'material':'','material_evidence':[]}]}
            for name,data in [('nwd_export_manifest.json',manifest),('nwd_geometry.json',[node]),
                ('nwd_sets.json',[]),('all_nodes.json',{'nodes':[{'has_geometry':True}]})]:
                (p/name).write_text(json.dumps(data),encoding='utf8')
            _,receipt=validate_export(p,'a'*64)
            self.assertTrue(receipt['local_coordinates_reference']['passed'])
            self.assertIn('nwd_local_triangles.bin',receipt['artifact_hashes'])
            self.assertFalse(receipt['release_allowed'])

if __name__=='__main__':unittest.main()
