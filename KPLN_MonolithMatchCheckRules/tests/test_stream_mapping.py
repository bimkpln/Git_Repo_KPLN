import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import json,tempfile,unittest
from pathlib import Path
from stream_mapping import entries
class StreamTests(unittest.TestCase):
    def check(self,text,expected):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory)/'mapping.json';p.write_text(text,encoding='utf8')
            for size in [1,2,7,1024]:self.assertEqual(dict(entries(p,size)),expected)
    def test_unicode_nested_and_exponents(self):
        d={'ключ':{'shape':'Линия\n"тест"','v':[1,2,3]},'a':1.2e30,'b':False,'c':None}
        self.check(json.dumps(d,ensure_ascii=False),d)
    def test_empty(self):self.check('{}',{})
    def test_truncated_value_fails(self):
        with self.assertRaises(ValueError):self.check('{"a":{"b":',{})
    def test_trailing_data_fails(self):
        with self.assertRaises(ValueError):self.check('{} 5',{})
if __name__=='__main__':unittest.main()
