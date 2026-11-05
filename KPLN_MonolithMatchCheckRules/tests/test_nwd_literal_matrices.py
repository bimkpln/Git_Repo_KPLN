import unittest,tempfile,hashlib,struct,json
from pathlib import Path
from nwd_literal_matrices import LiteralMatrices,NAME
from nwd_local_coordinates import callback_point

M=[-0.8461476059944555,0.5329486174762552,0,0,0.5329486174762552,0.8461476059944555,0,0,0,0,-1,0,7212003.4914964195,1041238.0544260767,38.06594487899552,1]

class LiteralMatrixTests(unittest.TestCase):
    def fixture(self,p):
        data=struct.pack('<ii16d',0,0,*M);(p/NAME).write_bytes(data)
        manifest={'literal_matrix_file':NAME,'literal_matrix_sha256':hashlib.sha256(data).hexdigest()}
        displayed=M.copy();displayed[12]=7212003.49149642
        return manifest,{'index':0,'fragmentRecords':[{'fragment_id':'0:0','local_to_world':displayed}]}
    def test_roundtrip_loss_in_JSON_does_not_change_literal_transform(self):
        with tempfile.TemporaryDirectory() as d:
            manifest,node=self.fixture(Path(d));reader=LiteralMatrices(d,manifest)
            try:
                row=reader.node(node);receipt=reader.finish()
                self.assertEqual(row['fragmentRecords'][0]['local_to_world'],M)
                self.assertEqual(receipt['JSON_display_matrices_with_differences'],1)
                self.assertEqual(callback_point((0.08202099800109863,-0.3034777045249939,0.4019028842449188),M)[0],7212003.260356525)
            finally:reader.close()
    def test_legacy_without_binary_is_unchanged(self):
        with tempfile.TemporaryDirectory() as d:
            reader=LiteralMatrices(d,{});node={'index':0};self.assertIs(reader.node(node),node);self.assertIsNone(reader.finish());reader.close()
    def test_changed_or_truncated_file_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);manifest,node=self.fixture(p);(p/NAME).write_bytes(b'wrong')
            with self.assertRaises(ValueError):LiteralMatrices(d,manifest)
    def test_undeclared_binary_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            self.fixture(Path(d))
            with self.assertRaises(ValueError):LiteralMatrices(d,{})
    def test_wrong_owner_or_order_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            manifest,node=self.fixture(Path(d));node['index']=1;reader=LiteralMatrices(d,manifest)
            try:
                with self.assertRaises(ValueError):reader.node(node)
            finally:reader.close()
    def test_unaccounted_record_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            manifest,node=self.fixture(Path(d));reader=LiteralMatrices(d,manifest)
            try:
                with self.assertRaises(ValueError):reader.finish()
            finally:reader.close()
    def test_escaping_path_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            manifest,node=self.fixture(Path(d));manifest['literal_matrix_file']='../outside'
            with self.assertRaises(ValueError):LiteralMatrices(d,manifest)
    def test_actual_import_uses_binary_before_recomposing_local_source(self):
        from nwd_import import validate_export
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);manifest,node=self.fixture(p)
            points=[(.08202099800109863,-.3034777045249939,.4019028842449188),(0.,0.,0.),(1.,1.,1.)]
            local=b''.join(struct.pack('<3d',*x) for x in points)
            (p/'nwd_local_triangles.bin').write_bytes(local)
            (p/'nwd_triangles.bin').write_bytes(b''.join(struct.pack('<3d',*callback_point(x,M)) for x in points))
            manifest.update(complete=True,schema=2,input_sha256='a'*64,geometry_nodes=1,scene_nodes=1,
                local_triangle_file='nwd_local_triangles.bin',local_triangle_sha256=hashlib.sha256(local).hexdigest())
            node.update(offset=0,triangles=1,fragments=1,excludedBySet=False,
                ancestry=[{'name':'instance','properties':[{'category':'Объект','name':'Id','value':'123'}]},{'name':'AR.nwc','properties':[]}])
            node['fragmentRecords'][0].update(offset=0,triangles=1,path_matches_node=True,excludedBySet=False,material='',material_evidence=[])
            for name,data in [('nwd_export_manifest.json',manifest),('nwd_geometry.json',[node]),('nwd_sets.json',[]),('all_nodes.json',{'nodes':[{'has_geometry':True}]})]:
                (p/name).write_text(json.dumps(data),encoding='utf8')
            fragments,receipt=validate_export(p,'a'*64)
            self.assertEqual(fragments[0]['local_to_world'],M)
            self.assertTrue(receipt['local_coordinates_reference']['passed'])
            self.assertIn(NAME,receipt['artifact_hashes'])
            self.assertFalse(receipt['release_allowed'])

if __name__=='__main__':unittest.main()
