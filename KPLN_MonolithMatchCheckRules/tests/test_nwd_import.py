import sys,unittest,tempfile,json
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from nwd_import import fragments_from_nodes,validate_export
from preflight import raw_fragments
from materials import classify_fragment

def node():
    matrix=[1.,0,0,0,0,1.,0,0,0,0,1.,0,0,0,0,1.]
    def p(name,v):return {'category':'Объект','name':name,'value':'DisplayString:'+v}
    return {'index':0,'offset':0,'triangles':2,'fragments':2,'excludedBySet':False,
        'ancestry':[{'name':'leaf','properties':[]},{'name':'instance','properties':[p('Id','123'),p('Категория','Стены')]},
                    {'name':'system','properties':[p('Id','456'),p('Категория','Лестницы')]},
                    {'name':'source.nwc','properties':[]}],
        'fragmentRecords':[{'fragment_id':'0:'+str(i),'offset':i*72,'triangles':1,'local_to_world':matrix.copy(),
            'material':m,'material_evidence':[{'value':m}],'path_matches_node':True,'excludedBySet':False}
            for i,m in enumerate(['Монолитный бетон','Сталь'])]}

class FragmentImportTests(unittest.TestCase):
    def test_separate_materials_and_nearest_instance(self):
        records=fragments_from_nodes([node()])
        self.assertEqual([r['material'] for r in records],['Монолитный бетон','Сталь'])
        self.assertTrue(all(r['element_id']=='123' and r['category']=='Стены' for r in records))
        self.assertEqual(raw_fragments([node()]),records)
    def test_legacy_preflight_cannot_bypass_fragment_guard(self):
        x=node();del x['fragmentRecords']
        with self.assertRaisesRegex(ValueError,'Legacy multi-fragment'):raw_fragments([x])
    def test_paint_excluded_even_with_approved_material(self):
        self.assertEqual(classify_fragment({'material':'КП_Краска (КРАС)'},['КП_Краска (КРАС)'])['decision'],'exclude')
    def test_legacy_aggregate_rejected(self):
        x=node();del x['fragmentRecords']
        with self.assertRaisesRegex(ValueError,'individual fragment'):fragments_from_nodes([x])
    def test_gap_rejected(self):
        x=node();x['fragmentRecords'][1]['offset']=144
        with self.assertRaisesRegex(ValueError,'byte range'):fragments_from_nodes([x])
    def test_misassigned_owner_rejected(self):
        x=node();x['fragmentRecords'][1]['path_matches_node']=False
        with self.assertRaisesRegex(ValueError,'owner path'):fragments_from_nodes([x])
    def test_nonfinite_matrix_rejected(self):
        x=node();x['fragmentRecords'][0]['local_to_world'][0]=float('nan')
        with self.assertRaisesRegex(ValueError,'transformation'):fragments_from_nodes([x])
    def test_singular_matrix_rejected(self):
        x=node();x['fragmentRecords'][0]['local_to_world'][0]=0
        with self.assertRaisesRegex(ValueError,'Singular'):fragments_from_nodes([x])
    def test_conflicting_material_rejected(self):
        x=node();x['fragmentRecords'][0]['material_evidence'][0]['value']='Сталь'
        with self.assertRaisesRegex(ValueError,'Conflicting'):fragments_from_nodes([x])
    def test_hide_membership_not_visibility(self):
        x=node();x['hidden']=True
        self.assertFalse(fragments_from_nodes([x])[0]['excluded_by_hide_set'])
        x['excludedBySet']=True
        for f in x['fragmentRecords']:f['excludedBySet']=True
        self.assertTrue(fragments_from_nodes([x])[0]['excluded_by_hide_set'])
    def test_incomplete_triangle_coverage_rejected(self):
        x=node();x['triangles']=3
        with self.assertRaisesRegex(ValueError,'coverage'):fragments_from_nodes([x])

class ExportAdmissionTests(unittest.TestCase):
    def fixture(self,p):
        def write(name,data):(p/name).write_text(json.dumps(data),encoding='utf8')
        write('nwd_export_manifest.json',{'complete':True,'schema':2,'input_sha256':'a'*64,'geometry_nodes':1,'scene_nodes':1})
        write('nwd_geometry.json',[node()]);write('nwd_sets.json',[])
        write('all_nodes.json',{'nodes':[{'has_geometry':True}]})
        (p/'nwd_triangles.bin').write_bytes(bytes(144))
    def test_valid_export_never_claims_geometry_check(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);self.fixture(p);f,r=validate_export(p,'a'*64)
            self.assertEqual(len(f),2);self.assertTrue(r['import_complete'])
            self.assertFalse(r['complete_geometry_check']);self.assertFalse(r['release_allowed'])
    def test_truncated_binary_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);self.fixture(p);(p/'nwd_triangles.bin').write_bytes(bytes(72))
            with self.assertRaisesRegex(ValueError,'Truncated'):validate_export(p,'a'*64)
    def test_wrong_model_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);self.fixture(p)
            with self.assertRaisesRegex(ValueError,'another NWD'):validate_export(p,'b'*64)
    def test_missing_non_geometric_scene_nodes_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);self.fixture(p);(p/'all_nodes.json').write_text('{"nodes":[]}',encoding='utf8')
            with self.assertRaisesRegex(ValueError,'scene tree'):validate_export(p,'a'*64)

if __name__=='__main__':unittest.main()
