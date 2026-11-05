import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import json,tempfile,hashlib,unittest
from nwd_metadata import validate_metadata_export
class MetadataTests(unittest.TestCase):
    def fixture(self,directory):
        r=[{'index':0,'parent':-1,'name':'fixture.nwd','model':'fixture.nwd','source':'fixture.nwd','properties':[],
            'bounds':None,'has_geometry':False,'hidden':False,'excluded_by_hide_set':False},
           {'index':1,'parent':0,'name':'AR.nwc','model':'fixture.nwd','source':'AR.nwc','properties':[],
            'bounds':[0,0,0,1,1,1],'has_geometry':True,'hidden':False,'excluded_by_hide_set':True}]
        p=Path(directory);(p/'nodes.json').write_text(json.dumps(r),encoding='utf8')
        m={'schema':1,'complete':True,'input':'fixture.nwd','input_sha256':'1'*64,'input_sha256_after_read':'1'*64,
            'reader_executable_sha256':'2'*64,'metadata_sha256':hashlib.sha256((p/'nodes.json').read_bytes()).hexdigest(),
            'geometry_extracted':False,'source_modified':False,'nodes':2,'excluded_set_nodes':1,
            'models':[{'file':'fixture.nwd','units':'Feet'}]}
        return p,m
    def check(self,change=None):
        with tempfile.TemporaryDirectory() as directory:
            p,m=self.fixture(directory)
            if change:change(p,m)
            (p/'manifest.json').write_text(json.dumps(m),encoding='utf8')
            return validate_metadata_export(p,'1'*64,'2'*64)
    def test_full_metadata_read_is_not_geometry_acceptance(self):
        p=self.check();self.assertEqual(p['scene_nodes'],2);self.assertFalse(p['complete_check']);self.assertFalse(p['geometry_check_performed'])
    def test_wrong_node_count_rejected(self):
        with self.assertRaises(ValueError):self.check(lambda p,m:m.update(nodes=1))
    def test_changed_nwd_rejected(self):
        with self.assertRaises(ValueError):self.check(lambda p,m:m.update(input_sha256_after_read='3'*64))
    def test_untrusted_reader_rejected(self):
        with self.assertRaises(ValueError):self.check(lambda p,m:m.update(reader_executable_sha256='3'*64))
    def test_geometry_export_is_not_metadata_contract(self):
        with self.assertRaises(ValueError):self.check(lambda p,m:m.update(geometry_extracted=True))
    def test_hide_membership_count_must_match(self):
        with self.assertRaises(ValueError):self.check(lambda p,m:m.update(excluded_set_nodes=0))
if __name__=='__main__':unittest.main()
