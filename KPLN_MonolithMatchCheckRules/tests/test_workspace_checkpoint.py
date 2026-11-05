import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import workspace_checkpoint as c

class CheckpointTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(); self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name);self.src=self.root/'source';self.src.mkdir()
        (self.src/'code').mkdir();(self.src/'code/a.txt').write_text('original')
        (self.src/'code/duplicate.txt').write_text('original')
        (self.src/'code/new.txt').write_text('second')
        self.cfg=self.root/'config.json';c.write_json(self.cfg,{'workspace_id':'test','include':['code']})
        self.remote=self.root/'remote'
        self.output=contextlib.redirect_stdout(io.StringIO());self.output.__enter__();self.addCleanup(self.output.__exit__,None,None,None)
    def build(self,name='cp1',baseline=None):
        cp=self.root/name;c.build(self.src,cp,self.cfg,baseline);return cp
    def test_roundtrip_duplicates_prefix_and_existing_target(self):
        cp=self.build();c.publish(cp,self.remote)
        c.restore(self.remote,'cp1',self.root/'restored')
        self.assertEqual((self.root/'restored/code/a.txt').read_text(),'original')
        self.assertEqual((self.root/'restored/code/duplicate.txt').read_text(),'original')
        c.restore(self.remote,'cp1',self.root/'partial',['code/new.txt'])
        self.assertFalse((self.root/'partial/code/a.txt').exists())
        with self.assertRaises(FileExistsError):c.restore(self.remote,'cp1',self.root/'restored')
    def test_incremental_and_deletion(self):
        cp=self.build();c.publish(cp,self.remote)
        (self.src/'code/new.txt').unlink();(self.src/'code/a.txt').write_text('modified')
        cp2=self.build('cp2',cp);m=c.read(cp2/'manifest.json');self.assertEqual(m['reused_files'],1)
        c.publish(cp2,self.remote);c.restore(self.remote,'cp2',self.root/'restored')
        self.assertFalse((self.root/'restored/code/new.txt').exists())
        self.assertEqual((self.root/'restored/code/a.txt').read_text(),'modified')
    def test_failed_upload_no_ready_then_retry(self):
        cp=self.build()
        with patch.object(c.shutil,'copyfile',side_effect=OSError('network offline')):
            with self.assertRaises(OSError):c.publish(cp,self.remote)
        self.assertFalse((self.remote/'checkpoints/cp1/READY.json').exists())
        c.publish(cp,self.remote);c.publish(cp,self.remote)
        self.assertTrue((self.remote/'checkpoints/cp1/READY.json').exists())
    def test_corrupt_existing_object_never_overwritten(self):
        cp=self.build();c.publish(cp,self.remote)
        obj=next((self.remote/'objects').rglob('*.zip'));obj.write_bytes(b'corrupt')
        with self.assertRaises(RuntimeError):c.publish(cp,self.remote)
        with self.assertRaises(RuntimeError):c.restore(self.remote,'cp1',self.root/'restore')
        self.assertEqual(obj.read_bytes(),b'corrupt')
    def test_unready_and_traversal(self):
        cp=self.build();c.publish(cp,self.remote)
        marker=self.remote/'checkpoints/cp1/READY.json';marker.unlink()
        with self.assertRaises(FileNotFoundError):c.restore(self.remote,'cp1',self.root/'restore')
        for name in ('../x','/root','C:/x','a\\b','a/../b'):
            with self.assertRaises(ValueError):c.relative(name)
    def test_modified_manifest_refused(self):
        cp=self.build();(cp/'manifest.json').write_text('{}')
        with self.assertRaises(RuntimeError):c.publish(cp,self.remote)
    def test_rehash_unchanged_timestamp(self):
        import os
        cp=self.build();c.publish(cp,self.remote)
        p=self.src/'code/a.txt';st=p.stat();p.write_text('changed!');os.utime(p,ns=(st.st_atime_ns,st.st_mtime_ns))
        cp2=self.build('cp2',cp)
        self.assertNotEqual(c.read(cp/'manifest.json')['files']['code/a.txt']['sha256'],c.read(cp2/'manifest.json')['files']['code/a.txt']['sha256'])

if __name__=='__main__':unittest.main()
