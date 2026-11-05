import datetime,sys,tempfile,unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from result_directory import folder_name,allocate
from policy import resolve

class ResultDirectoryTests(unittest.TestCase):
    def cfg(self,**kwargs):return dict(project='p',stage='RD',building_id='K1',result_scope_label='Блок1',**kwargs)
    def date(self):return datetime.datetime(2026,10,7,10,tzinfo=datetime.timezone.utc)
    def test_company_name(self):
        self.assertEqual(folder_name(self.cfg(),self.date()),'2026.10.07_Монолит_Фулл_Блок1')
    def test_moscow_date_not_host_date(self):
        t=datetime.datetime(2026,10,6,22,tzinfo=datetime.timezone.utc)
        self.assertEqual(folder_name(self.cfg(),t),'2026.10.07_Монолит_Фулл_Блок1')
    def test_lite_requires_exclusions(self):
        with self.assertRaises(ValueError):folder_name(self.cfg(check_variant='lite'),self.date())
        with self.assertRaises(ValueError):folder_name(self.cfg(additional_exclusions=['Без отверстий']),self.date())
    def test_mode_changes_profile_hash(self):
        a=resolve(self.cfg());b=resolve(self.cfg(check_variant='lite',additional_exclusions=['Без инженерных отверстий']))
        self.assertNotEqual(a['profile_hash'],b['profile_hash'])
        self.assertEqual(b['additional_exclusions'],['Без инженерных отверстий'])
    def test_previous_results_not_overwritten_concurrently(self):
        with tempfile.TemporaryDirectory() as folder:
            first=allocate(folder,self.cfg(),self.date());(first/'keep.txt').write_text('original')
            with ThreadPoolExecutor(max_workers=4) as pool:
                targets=list(pool.map(lambda _:allocate(folder,self.cfg(),self.date()),range(4)))
            self.assertEqual(len(set(targets)),4);self.assertNotIn(first,targets)
            self.assertEqual((first/'keep.txt').read_text(),'original')
            self.assertIn('2026.10.07_Монолит_Фулл_Блок1_02',[p.name for p in targets])
    def test_scope_cannot_escape_root(self):
        c=self.cfg();c['result_scope_label']='../other'
        with self.assertRaises(ValueError):folder_name(c,self.date())
