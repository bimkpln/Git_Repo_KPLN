import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import json,tempfile,unittest
from stream_sequence import items
class ArrayStreamTests(unittest.TestCase):
    def check(self,text,expected):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory)/'a.json';p.write_text(text,encoding='utf8')
            for size in [1,2,9,1024]:self.assertEqual(list(items(p,size)),expected)
    def test_unicode_nested_values_and_exponent(self):
        x=[{'имя':'Дверь\n"А"','p':[1,2]},1.2e30,False,None];self.check(json.dumps(x,ensure_ascii=False),x)
    def test_empty(self):self.check('[]',[])
    def test_truncated(self):
        with self.assertRaises(ValueError):self.check('[{"a":',[])
    def test_trailing_data_rejected(self):
        with self.assertRaises(ValueError):self.check('[]{}',[])
if __name__=='__main__':unittest.main()
