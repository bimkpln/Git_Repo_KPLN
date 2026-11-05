import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from preflight import prepare
from policy import digest

class FragmentAssignmentsTests(unittest.TestCase):
    def fixture(self):
        scope={'project':'test','stage':'RD','building_id':'B1','section_id':None}
        f={'fragment_id':'0:0','material':'','category':'Обобщенные модели'}
        cfg={**scope,'input_nwd_sha256':'a'*64,'verified_fragment_assignments':[{
            'fragment_id':'0:0','input_nwd_sha256':'a'*64,'fragment_metadata_sha256':digest(f),
            'values':{'material':'Монолитный железобетон','construction_method':'монолит'},
            'evidence':{'type':'user_confirmation','statement':'Confirmed cast concrete stair','registry_scope':scope}}]}
        return cfg,f
    def test_explicit_assignment_retains_evidence_and_changes_identity(self):
        cfg,f=self.fixture();before=prepare({k:v for k,v in cfg.items() if k!='verified_fragment_assignments'},[f])
        after=prepare(cfg,[f]);self.assertTrue(after['ready_for_geometry'])
        self.assertNotEqual(before['fragment_inventory_sha256'],after['fragment_inventory_sha256'])
        self.assertEqual(f['material'],'')
    def test_another_model_rejected(self):
        cfg,f=self.fixture();cfg['input_nwd_sha256']='b'*64
        with self.assertRaises(ValueError):prepare(cfg,[f])
    def test_changed_source_fragment_rejected(self):
        cfg,f=self.fixture();f['category']='Стены'
        with self.assertRaises(ValueError):prepare(cfg,[f])
    def test_another_building_rejected(self):
        cfg,f=self.fixture();cfg['building_id']='B2'
        with self.assertRaises(ValueError):prepare(cfg,[f])
    def test_assignment_cannot_change_hide_membership(self):
        cfg,f=self.fixture();cfg['verified_fragment_assignments'][0]['values']['excluded_by_hide_set']=False
        with self.assertRaises(ValueError):prepare(cfg,[f])

if __name__=='__main__':unittest.main()
