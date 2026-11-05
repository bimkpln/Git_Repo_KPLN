import hashlib,json,subprocess,sys,tempfile,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]

class PrepareTests(unittest.TestCase):
    def run_prepare(self,p,extra=None,use_root=False):
        cfg={'project':'test','stage':'RD','building_id':'K1','section_id':'S1',**(extra or {})}
        (p/'config.json').write_text(json.dumps(cfg),encoding='utf8');(p/'model.nwd').write_bytes(b'synthetic NWD placeholder - no geometry check')
        return subprocess.run([sys.executable,'-X','utf8','-B',str(ROOT/'scripts/prepare-run.py'),'--config',str(p/'config.json'),
            '--nwd',str(p/'model.nwd'),'--results-root' if use_root else '--output',str(p/('reports' if use_root else 'run')),'--operator','test'],capture_output=True,text=True,encoding='utf8')
    def test_automatic_folder_saved_in_run_and_repeat_separate(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);cfg={'result_scope_label':'К2С14','check_variant':'lite','additional_exclusions':['Без инженерных отверстий']}
            paths=[]
            for _ in range(2):
                r=self.run_prepare(p,cfg,use_root=True);self.assertEqual(r.returncode,0,r.stderr)
                path=Path(json.loads(r.stdout)['run']);paths.append(path)
                run=json.loads(path.read_text(encoding='utf8'))
                self.assertEqual(run['result_directory'],str(path.parent))
                self.assertEqual(run['profile']['additional_exclusions'],cfg['additional_exclusions'])
                self.assertIn('_Монолит_Лайт_К2С14',path.parent.name)
            self.assertNotEqual(paths[0],paths[1]);self.assertTrue(paths[1].parent.name.endswith('_02'))
    def test_snapshot_and_no_claim_of_geometry_check(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);r=self.run_prepare(p);self.assertEqual(r.returncode,0,r.stderr)
            run=json.loads((p/'run/run.json').read_text(encoding='utf8'));self.assertEqual(run['status'],'prepared_not_checked')
            self.assertEqual(run['profile']['general_mm'],15)
            manifest=json.loads((p/'run/rules_manifest.json').read_text())
            self.assertTrue(all(hashlib.sha256((p/'run/rules_snapshot'/n).read_bytes()).hexdigest()==h for n,h in manifest.items()))
            self.assertFalse(any(n.startswith('local/') for n in manifest))
            self.assertIn('engine/NwdFragmentProbe.cs',manifest)
            self.assertEqual(run['input_nwd_sha256'],hashlib.sha256((p/'model.nwd').read_bytes()).hexdigest())
            self.assertNotEqual(self.run_prepare(p).returncode,0)
    def test_pd_cannot_start_without_documented_scope(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);r=self.run_prepare(p,{'stage':'PD'})
            self.assertNotEqual(r.returncode,0);self.assertFalse((p/'run/run.json').exists())
    def test_fifty_confirmation_is_in_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);r=self.run_prepare(p,{'fully_constructed':True,'construction_confirmation':'User confirmed this run'})
            self.assertEqual(r.returncode,0,r.stderr)
            run=json.loads((p/'run/run.json').read_text(encoding='utf8'));self.assertEqual(run['profile']['general_mm'],50)
            self.assertEqual(run['profile']['construction_confirmation'],'User confirmed this run')
