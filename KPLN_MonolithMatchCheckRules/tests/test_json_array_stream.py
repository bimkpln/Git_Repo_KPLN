import unittest,json,tempfile
from pathlib import Path
from json_array_stream import array_items

class ArrayStreamTests(unittest.TestCase):
    def read(self,text,key=None,chunk=3):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'x.json';p.write_text(text,encoding='utf-8-sig')
            return list(array_items(p,key,chunk))
    def test_objects_strings_and_unicode_across_one_character_chunks(self):
        data=[{'имя':'а\\б"в','values':[1,2,None,True]},'x',{},[]]
        self.assertEqual(self.read(json.dumps(data,ensure_ascii=False),chunk=1),data)
    def test_wrapper_and_empty_array(self):
        self.assertEqual(self.read('{"nodes":[]}','nodes'),[])
    def test_large_single_record(self):
        data=[{'text':'x'*20000},{'last':True}]
        self.assertEqual(self.read(json.dumps(data),chunk=17),data)
    def test_numbers_cannot_be_truncated_at_chunk_boundary(self):
        self.assertEqual(self.read('[123456789,-12.5,1.23e-20]',chunk=1),[123456789,-12.5,1.23e-20])
    def test_wrong_wrapper_is_rejected(self):
        with self.assertRaises(ValueError):self.read('{"wrong":[]}','nodes')
    def test_duplicate_or_additional_wrapper_keys_are_rejected(self):
        for text in ['{"nodes":[],"nodes":[]}','{"nodes":[],"extra":1}']:
            with self.subTest(text=text),self.assertRaises(ValueError):self.read(text,'nodes')
    def test_malformed_and_trailing_data_are_rejected(self):
        for text in ['[1,]','[1 2]','[1','{}','[tru]','[]x','[1:2]','["unterminated]']:
            with self.subTest(text=text),self.assertRaises(ValueError):self.read(text)
    def test_chunk_size_must_be_positive(self):
        for chunk in [0,-1,False,1.5]:
            with self.subTest(chunk=chunk),self.assertRaises(ValueError):self.read('[]',chunk=chunk)

if __name__=='__main__':unittest.main()
