import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from semantic_evidence import declared_role_candidates
class SemanticTests(unittest.TestCase):
    def test_concrete_primer_and_completion_are_not_windows(self):
        for name in ['бетоноконтакт','Окончание стены']:
            self.assertFalse(declared_role_candidates({'name':name})['door_window'])
    def test_family_separators_keep_declared_window_names(self):
        self.assertTrue(declared_role_candidates({'name':'123_Оконный_Блок'})['door_window'])
    def test_workset_and_boolean_parameter_name_do_not_prove_mounting(self):
        node={'name':'Перекрытие','properties':[{'category':'Объект','name':'Рабочий набор','value':'DisplayString:Монтажные отверстия'},
            {'category':'Объект','name':'Монтажное отверстие','value':'Boolean:False'}]}
        self.assertFalse(declared_role_candidates(node)['mounting'])
    def test_explicit_quantity_role_is_preserved(self):
        node={'name':'Стена','properties':[{'category':'Объект','name':'Орг.ГруппаВедомостьРасходаБетона','value':'DisplayString:Приямок монолитный'}]}
        evidence=declared_role_candidates(node)['pit'];self.assertEqual(evidence[0]['property'],'Орг.ГруппаВедомостьРасходаБетона')
if __name__=='__main__':unittest.main()
