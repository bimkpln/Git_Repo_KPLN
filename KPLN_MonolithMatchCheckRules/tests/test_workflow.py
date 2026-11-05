import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import copy,json,tempfile,unittest
import zipfile
from pathlib import Path
from policy import resolve,opening_decision,clip_triangle_to_views
from materials import classify_fragment,select_fragments
from preflight import prepare,validate_engine_receipt
from history import seed,reconcile,export_history
from update_history import prepare as prepare_history,commit,filehash
from geometry_identity import bind_geometry
from registry_scope import registry_path

def config(stage='RD'):
    return {'project':'p','stage':stage,'building_id':'K2','section_id':'S14','input_nwd_sha256':'a'*64}
def view():
    return {'id':'plan1','source':'plan.pdf','source_sha256':'b'*64,'registration_evidence':'approved transform and level depths',
        'volume':{'type':'oriented_box','origin_mm':[0,0,0],'basis':[[1,0,0],[0,1,0],[0,0,1]],'bounds_mm':[[0,0,0],[10,10,10]]}}
def issue(n='M-001',x=0,owner='AR|1'):
    return {'number':n,'source':owner,'category':'01','measurement_rule':'surface','measurement_mm':30,
        'identity_geometry_sha256':'d'*64,
        'bounds_mm':[[x,0,0],[x+100,100,100]],'children':[{'number':n,'AR':owner,'KR':'KR|1','point_mm':[x,0,0],'world_m':[x/1000,0,0]}]}
def coverage(reg):
    return {'project':'p','stage':'PD','building_id':'K2','section_id':'S14','complete':True,'material_filter_complete':True,
        'geometry_receipt_sha256':'c'*64,'same_rules':True,'fully_rechecked_issue_uids':list(reg['issues']),'number_prefix':'M-'}

class WorkflowTests(unittest.TestCase):
    def test_pd_requires_drawings(self):
        with self.assertRaises(ValueError):resolve(config('PD'))
    def test_pd_full_only_explicit_test(self):
        c=config('PD')|{'scope':'full_model'}
        with self.assertRaises(ValueError):resolve(c)
        self.assertEqual(resolve(c|{'explicit_full_model_test':True})['scope'],'full_model')
    def test_rd_full_only(self):
        self.assertEqual(resolve(config())['scope'],'full_model')
        with self.assertRaises(ValueError):resolve(config()|{'scope':'documented_views','views':[view()]})
    def test_rd_hole_changes_threshold_not_edge_metric(self):
        rd=resolve(config());pd=resolve(config('PD')|{'views':[view()]})
        self.assertTrue(opening_decision(rd,'technical',[4,4,4,4])['issue'])
        self.assertFalse(opening_decision(pd,'technical',[4,4,4,4])['issue'])
    def test_door_opposite_pairs_and_equality(self):
        p=resolve(config())
        self.assertFalse(opening_decision(p,'door',[8,7,8,7])['issue'])
        self.assertTrue(opening_decision(p,'window',[8,8,0,0])['issue'])
    def test_missing_has_no_size_tolerance(self):
        self.assertTrue(opening_decision(resolve(config()),'technical',missing=True)['issue'])
    def test_rd_does_not_inherit_pd_exclusions(self):
        self.assertNotIn('roof_shaft_cap',resolve(config())['exclude'])
        self.assertNotIn('MEP_equipment_foundation',resolve(config())['exclude'])
    def test_fifty_not_implicit(self):
        with self.assertRaises(ValueError):resolve(config()|{'fully_constructed':True})
        self.assertEqual(resolve(config())['general_mm'],15)
    def test_explicit_building_confirmation_activates_fifty(self):
        p=resolve(config()|{'fully_constructed':True,'construction_confirmation':'User explicitly confirmed this building for this run'})
        self.assertEqual([p[k] for k in ['general_mm','flight_underside_vertical_mm','technical_hole_sum_four_mm','door_window_opposite_pair_mm']],[50]*4)
        self.assertTrue(opening_decision(p,'door',missing=True)['issue'])
        self.assertFalse(opening_decision(p,'technical',[20,10,10,10])['issue'])
        self.assertTrue(opening_decision(p,'technical',[21,10,10,10])['issue'])
    def test_fifty_never_applies_to_pd(self):
        with self.assertRaises(ValueError):resolve(config('PD')|{'views':[view()],'fully_constructed':True,'construction_confirmation':'user'})
    def test_crop_not_centroid_test(self):
        p=resolve(config('PD')|{'views':[view()]})
        tri=[[-100,2,5],[2,2,5],[2,3,5]]
        result=clip_triangle_to_views(tri,p);self.assertTrue(result)
        self.assertTrue(all(0<=a[0]<=10 for a in result[0]['polygon_mm']))
    def test_outside_scope_empty(self):
        p=resolve(config('PD')|{'views':[view()]})
        self.assertEqual(clip_triangle_to_views([[0,0,20],[1,0,20],[0,1,20]],p),[])
    def test_invalid_crop_basis_rejected(self):
        v=view();v['volume']['basis'][0]=[2,0,0]
        with self.assertRaises(ValueError):resolve(config('PD')|{'views':[v]})
    def test_hide_set_overrides_concrete(self):
        self.assertEqual(classify_fragment({'material':'Монолитный железобетон','excluded_by_hide_set':True})['decision'],'exclude')
    def test_concrete_without_method_unknown(self):
        self.assertEqual(classify_fragment({'material':'Железобетон','category':'Стены'})['decision'],'unresolved')
    def test_precast_concrete_excluded(self):
        self.assertEqual(classify_fragment({'material':'Железобетон','construction_method':'сборный'})['decision'],'exclude')
    def test_mixed_element_fragment_selection(self):
        items=[{'fragment_id':str(i),'material':m} for i,m in enumerate(['Монолитный железобетон','Утеплитель','Сталь'])]
        self.assertEqual(select_fragments(items)['counts'],{'include':1,'exclude':2,'unresolved':0})
    def test_unknown_blocks_engine(self):
        p=prepare(config(),[{'fragment_id':'1','material':'unknown'}])
        self.assertFalse(p['ready_for_geometry'])
        with self.assertRaises(ValueError):validate_engine_receipt(p,{})
    def test_legacy_engine_cannot_claim_new_profile(self):
        p=prepare(config(),[{'fragment_id':'1','material':'Монолитный бетон'}])
        with self.assertRaises(ValueError):validate_engine_receipt(p,{'complete':True})
    def baseline(self):return seed([issue()],'p','PD','v1','2026-10-22','a'*64,building_id='K2',section_id='S14')
    def test_unchanged_number_date_and_marker(self):
        reg=self.baseline();new=issue('temporary99');r=reconcile(reg,[new],'v2','2026-10-23',coverage(reg));self.assertTrue(r['complete'])
        entries,_=export_history([new],r['registry'],r['physical_uid_map'])
        self.assertEqual((entries[0]['number'],entries[0]['children'][0]['number'],entries[0]['first_detected'],entries[0]['lifecycle_status']),('M-001','M-001','2026-10-22','unworked'))
    def test_new_issue_number_never_reuses_existing(self):
        reg=self.baseline();new=issue('M-001',500,owner='AR|2');r=reconcile(reg,[new],'v2','2026-10-23',coverage(reg));e,_=export_history([new],r['registry'],r['physical_uid_map'])
        self.assertEqual(e[0]['number'],'M-002');self.assertEqual(e[0]['children'][0]['number'],'M-002');self.assertEqual(e[0]['lifecycle_status'],'new')
    def test_missing_only_resolved_when_fully_checked(self):
        reg=self.baseline();uid=next(iter(reg['issues']))
        c=coverage(reg);r=reconcile(reg,[],'v2','2026-10-23',c);self.assertEqual(r['registry']['issues'][uid]['status'],'resolved')
        c['fully_rechecked_issue_uids']=[];r=reconcile(reg,[],'v2','2026-10-23',c);self.assertEqual(r['registry']['issues'][uid]['status'],'not_checked')
    def test_changed_tolerance_cannot_close(self):
        reg=self.baseline();c=coverage(reg)|{'same_rules':False};r=reconcile(reg,[],'v2','2026-10-23',c)
        self.assertEqual(next(iter(r['registry']['issues'].values()))['status'],'not_checked')
    def test_changed_geometry_blocks_guessing(self):
        reg=self.baseline();new=issue(x=3);r=reconcile(reg,[new],'v2','2026-10-23',coverage(reg))
        self.assertFalse(r['complete']);self.assertIsNone(r['registry'])
    def test_cross_stage_rejected(self):
        reg=self.baseline()
        with self.assertRaises(ValueError):reconcile(reg,[],'v2','2026-10-23',coverage(reg)|{'stage':'RD'})
    def test_cross_building_or_section_rejected(self):
        reg=self.baseline()
        for changed in [{'building_id':'K3'},{'section_id':'S15'},{'section_id':None}]:
            with self.assertRaises(ValueError):reconcile(reg,[],'v2','2026-10-23',coverage(reg)|changed)
    def test_registry_path_has_project_stage_building_section(self):
        self.assertEqual(registry_path('registries',config('PD')).as_posix(),'registries/p/PD/K2/S14/registry.json')
        self.assertNotEqual(registry_path('registries',config()),registry_path('registries',config('PD')))
    def test_retired_registry_cannot_receive_observations(self):
        reg=self.baseline()|{'superseded_by':'new/registry.json'}
        with self.assertRaises(ValueError):reconcile(reg,[],'v2','2026-10-23',coverage(reg))
    def test_duplicate_run_rejected(self):
        reg=self.baseline()
        with self.assertRaises(ValueError):reconcile(reg,[],'v1','2026-10-23',coverage(reg))
    def test_reopened_preserves_original_number(self):
        reg=self.baseline();r=reconcile(reg,[],'v2','2026-10-23',coverage(reg));reg=r['registry']
        r=reconcile(reg,[issue()],'v3','2026-10-24',coverage(reg));self.assertEqual(next(iter(r['registry']['issues'].values()))['status'],'reopened')
    def test_reopened_after_out_of_scope_preserves_resolution_event(self):
        reg=self.baseline();reg=reconcile(reg,[],'v2','2026-10-23',coverage(reg))['registry']
        reg=reconcile(reg,[],'v3','2026-10-24',coverage(reg)|{'fully_rechecked_issue_uids':[]})['registry']
        reg=reconcile(reg,[issue()],'v4','2026-10-25',coverage(reg))['registry']
        v=next(iter(reg['issues'].values()))
        self.assertEqual(v['status'],'reopened')
        self.assertEqual([e['status'] for e in v['status_events']],['baseline','resolved','not_checked','reopened'])
    def test_same_bbox_different_geometry_not_exact_match(self):
        reg=self.baseline();g=issue()|{'identity_geometry_sha256':'e'*64}
        self.assertFalse(reconcile(reg,[g],'v2','2026-10-23',coverage(reg))['complete'])
    def test_history_without_geometry_fingerprint_rejected(self):
        g=issue();del g['identity_geometry_sha256']
        with self.assertRaises(ValueError):seed([g],'p','PD','v1','2026-10-22','a'*64,building_id='K2',section_id='S14')
    def test_geometry_fingerprint_ignores_paths_but_detects_contents(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);g=issue()|{'certified_meshes':{'AR':['a.npz']}}
            for n in ['a.npz','b.npz']:
                with zipfile.ZipFile(p/n,'w') as z:z.writestr('vertices.npy',b'geometry1')
            one=bind_geometry([g],p)[0]['identity_geometry_sha256']
            g['certified_meshes']['AR']=['b.npz']
            self.assertEqual(one,bind_geometry([g],p)[0]['identity_geometry_sha256'])
            with zipfile.ZipFile(p/'b.npz','w') as z:z.writestr('vertices.npy',b'geometry2')
            self.assertNotEqual(one,bind_geometry([g],p)[0]['identity_geometry_sha256'])
    def test_history_commit_requires_both_verified_reports(self):
        with tempfile.TemporaryDirectory() as directory:
            p=Path(directory);reg=self.baseline()
            for name,data in [('registry.json',reg),('entries.json',[issue()]),('coverage.json',coverage(reg))]:
                (p/name).write_text(json.dumps(data),encoding='utf8')
            prepare_history(p/'registry.json',p/'entries.json',p/'coverage.json',p/'out','v2','2026-10-23')
            before=filehash(p/'registry.json');t=json.loads((p/'out/transaction.json').read_text())
            (p/'report.html').write_text('verified report',encoding='utf8')
            receipt={'passed':True,'canonical_entries_sha256':t['entries_sha256'],'artifact_hashes':{'report.html':filehash(p/'report.html')}}
            (p/'good.json').write_text(json.dumps(receipt));(p/'bad.json').write_text(json.dumps(receipt|{'passed':False}))
            with self.assertRaises(ValueError):commit(p/'out',p/'good.json',p/'bad.json')
            self.assertEqual(filehash(p/'registry.json'),before)
            commit(p/'out',p/'good.json',p/'good.json')
            self.assertNotEqual(filehash(p/'registry.json'),before)
            self.assertTrue((p/'backups'/before/'registry.json').is_file())
    def test_stale_closure_certificate_rejected(self):
        c=config()|{'verified_original_closures':[{'fragment_id':'1','input_nwd_sha256':'x','fragment_metadata_sha256':'bad','proof_sha256':'b'*64}]}
        with self.assertRaises(ValueError):prepare(c,[{'fragment_id':'1','material':''}])
