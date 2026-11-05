import os,subprocess,tempfile,unittest
from pathlib import Path
from urllib.parse import urlparse,parse_qs
ROOT=Path(__file__).resolve().parents[1]
PS=Path(os.environ.get('SystemRoot','C:/Windows'))/'System32/WindowsPowerShell/v1.0/powershell.exe'

@unittest.skipUnless(os.name=='nt' and PS.is_file(),'Windows launcher')
class LauncherTests(unittest.TestCase):
    def run_ps(self,args):
        return subprocess.run([str(PS),'-NoProfile','-NonInteractive','-File',*map(str,args)],capture_output=True,timeout=20)
    def test_preview_preserves_prompt_and_workspace(self):
        with tempfile.TemporaryDirectory(prefix='monolith ') as folder:
            r=self.run_ps([ROOT/'scripts/new-check.ps1','-Workspace',folder,'-PreviewOnly'])
            self.assertEqual(r.returncode,0,r.stderr)
            uri=urlparse(r.stdout.decode('ascii').strip());self.assertEqual((uri.scheme,uri.netloc),('codex','new'))
            q=parse_qs(uri.query);self.assertEqual(set(q),{'prompt','path'})
            self.assertEqual(q['prompt'][0],(ROOT/'examples/start-request.txt').read_bytes().decode('utf8'))
            self.assertEqual(Path(q['path'][0]),Path(folder))
    def test_launch_branch_only_invokes_uri_handler(self):
        with tempfile.TemporaryDirectory() as folder:
            r=self.run_ps([ROOT/'tests/launcher-harness.ps1','-Launcher',ROOT/'scripts/new-check.ps1','-Workspace',folder])
            self.assertEqual(r.returncode,0,r.stderr)
            self.assertTrue(r.stdout.decode('ascii').strip().startswith('MOCK_URI:codex://new?prompt='))
    def test_relative_workspace_rejected(self):
        r=self.run_ps([ROOT/'scripts/new-check.ps1','-Workspace','.','-PreviewOnly'])
        self.assertNotEqual(r.returncode,0)
        self.assertNotIn(b'codex://new',r.stdout)
