import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from nwd_scene_nodes import scene_nodes
class SceneTests(unittest.TestCase):
    def sample(self):return {'nodes':[{'scene_index':0,'parent_index':-1,'properties':[]},{'scene_index':1,'parent_index':0,'properties':[]}]}
    def test_envelope_is_unwrapped_and_every_node_counted(self):
        p=self.sample();self.assertEqual(len(p),1);self.assertEqual(len(scene_nodes(p,2)),2)
    def test_wrong_manifest_count_rejected(self):
        with self.assertRaises(ValueError):scene_nodes(self.sample(),1)
    def test_wrong_root_rejected(self):
        with self.assertRaises(ValueError):scene_nodes(self.sample()['nodes'],2)
    def test_parent_cycle_rejected(self):
        p=self.sample();p['nodes'][1]['parent_index']=1
        with self.assertRaises(ValueError):scene_nodes(p,2)
    def test_boolean_index_and_count_rejected(self):
        p=self.sample();p['nodes'][0]['scene_index']=False
        with self.assertRaises(ValueError):scene_nodes(p,2)
        with self.assertRaises(ValueError):scene_nodes(self.sample(),True)
if __name__=='__main__':unittest.main()
