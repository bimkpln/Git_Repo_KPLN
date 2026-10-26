"""Synthetic behavioral tests: scope isolation and explicit override semantics."""
import json,tempfile,unittest
from pathlib import Path
from resolve_rules import resolve

class ResolverTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
        self.put('VERSION','fixture',False)
        self.put('registry.json',dict(schema_version=1,departments={'OV':{'path':'departments/OV/rules.json'},'AR':{'path':'departments/AR/rules.json'}},projects={'P':{'departments':{'OV':{'path':'projects/P/departments/OV/rules.json'}}}}))
        self.base=dict(id='R',title='test',active=True,condition='condition',disposition='review',severity='critical',source='fixture')
        self.layer('rules/common.json',{'level':'common'},[self.base])
        self.layer('departments/OV/rules.json',{'level':'department','department':'OV'},[dict(self.base,overrides=True,reason='department decision',severity='medium')])
        self.layer('departments/AR/rules.json',{'level':'department','department':'AR'},[])
        self.layer('projects/P/departments/OV/rules.json',{'level':'project_department','project':'P','department':'OV'},[dict(self.base,overrides=True,reason='project decision',severity='warning')])
    def tearDown(self):self.tmp.cleanup()
    def put(self,path,value,is_json=True):
        p=self.root/path;p.parent.mkdir(parents=True,exist_ok=True);p.write_text(json.dumps(value) if is_json else value,encoding='utf-8')
    def layer(self,path,scope,rules):self.put(path,dict(schema_version=1,scope=scope,rules=rules))
    def test_specificity(self):
        r=resolve(self.root,'OV','P');self.assertEqual(r['rules'][0]['severity'],'warning');self.assertEqual(len(r['history']['R']),3)
    def test_department_only(self):self.assertEqual(resolve(self.root,'OV')['rules'][0]['severity'],'medium')
    def test_no_cross_department_leak(self):self.assertEqual(resolve(self.root,'AR','P')['rules'][0]['severity'],'critical')
    def test_unknown_context(self):
        r=resolve(self.root,'OTHER','OTHER');self.assertEqual(len(r['layers']),1);self.assertTrue(r['warnings'])
    def test_missing_department(self):self.assertEqual(len(resolve(self.root,project='P')['layers']),1)
    def test_override_requires_reason(self):
        self.layer('departments/OV/rules.json',{'level':'department','department':'OV'},[dict(self.base,overrides=True)])
        with self.assertRaises(ValueError):resolve(self.root,'OV')
    def test_implicit_override_rejected(self):
        self.layer('departments/OV/rules.json',{'level':'department','department':'OV'},[self.base])
        with self.assertRaises(ValueError):resolve(self.root,'OV')
    def test_duplicate_rejected(self):
        self.layer('rules/common.json',{'level':'common'},[self.base,self.base])
        with self.assertRaises(ValueError):resolve(self.root)
    def test_disable_and_reenable(self):
        self.layer('departments/OV/rules.json',{'level':'department','department':'OV'},[dict(self.base,active=False,overrides=True,reason='disable')])
        self.assertFalse(resolve(self.root,'OV')['rules'][0]['active']);self.assertTrue(resolve(self.root,'OV','P')['rules'][0]['active'])
    def test_path_escape_rejected(self):
        self.put('registry.json',dict(schema_version=1,departments={'OV':{'path':'../outside.json'}},projects={}))
        with self.assertRaises(ValueError):resolve(self.root,'OV')

    def test_common_only_preserves_context_without_local_rules(self):
        r=resolve(self.root,'OV','P',common_only=True)
        self.assertEqual(r['rules'][0]['severity'],'critical')
        self.assertEqual(r['context'],dict(department='OV',project='P'))
        self.assertEqual(r['selection_mode'],'common_only')
        self.assertEqual([x['path'] for x in r['layers']],['rules/common.json'])

    def test_common_only_does_not_read_local_layers(self):
        self.put('departments/OV/rules.json','invalid json',False)
        self.assertEqual(len(resolve(self.root,'OV','P',common_only=True)['layers']),1)
        with self.assertRaises(ValueError):resolve(self.root,'OV','P')

    def test_no_cross_project_leak(self):
        r=resolve(self.root,'OV','OTHER')
        self.assertEqual(r['rules'][0]['severity'],'medium')
        self.assertEqual(len(r['layers']),2)

    def test_full_record_replacement(self):
        self.layer('rules/common.json',{'level':'common'},[dict(self.base,extra='must not survive')])
        self.assertNotIn('extra',resolve(self.root,'OV')['rules'][0])

    def test_unknown_override_rejected(self):
        self.layer('departments/AR/rules.json',{'level':'department','department':'AR'},[dict(self.base,id='ABSENT',overrides=True,reason='test')])
        with self.assertRaises(ValueError):resolve(self.root,'AR')

    def test_wrong_scope_rejected(self):
        self.layer('departments/AR/rules.json',{'level':'department','department':'OV'},[])
        with self.assertRaises(ValueError):resolve(self.root,'AR')

    def test_override_reason_must_be_nonempty_text(self):
        for reason in (None,False,10,[],{},'  '):
            with self.subTest(reason=reason):
                self.layer('departments/OV/rules.json',{'level':'department','department':'OV'},[dict(self.base,overrides=True,reason=reason)])
                with self.assertRaises(ValueError):resolve(self.root,'OV')

    def test_disable_requires_reason(self):
        self.layer('rules/common.json',{'level':'common'},[dict(self.base,active=False,reason=None)])
        with self.assertRaises(ValueError):resolve(self.root)

    def test_deleting_override_restores_inheritance(self):
        self.layer('projects/P/departments/OV/rules.json',{'level':'project_department','project':'P','department':'OV'},[])
        self.assertEqual(resolve(self.root,'OV','P')['rules'][0]['severity'],'medium')

if __name__=='__main__':unittest.main()
