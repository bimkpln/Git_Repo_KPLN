import sys,unittest,copy
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from policy import resolve,opening_decision
from project_opening_rules import evaluate
from preflight import validate_engine_receipt

class SgmProjectOpeningTests(unittest.TestCase):
    def setUp(self):
        self.cfg={'project':'Сочи гостиница Москва','stage':'RD','building_id':'Блок1','section_id':None,'scope':'full_model'}
        self.profile=resolve(self.cfg)
        self.e={'opening_present_AR':True,'opening_present_KR':True,'same_physical_opening_proved':True,
                'AR_main_contour_contained_in_KR_over_whole_depth':True,'all_AR_internal_dividers_examined':True,
                'geometry_evidence_sha256':'a'*64,'KR_outward_side_intervals_mm':[[0,0],[200,200],[10,10],[30,30]],
                'AR_lintel_width_intervals_mm':[[300,300]]}
    def result(self):return opening_decision(self.profile,'technical',[0,200,10,30],project_geometry_evidence=self.e,host_kind='floor')
    def test_only_SGM_gets_rule(self):
        for project in ['СГРВН','sgrvn','Другой проект','СГМ']:
            p=resolve({**self.cfg,'project':project});self.assertNotIn('project_opening_rules',p)
    def test_all_sections_of_registered_project(self):
        p=resolve({**self.cfg,'building_id':'Блок2','section_id':'Секция7'});self.assertIn('project_opening_rules',p)
    def test_project_rule_is_in_profile_hash(self):
        p=copy.deepcopy(self.profile);p['project_opening_rules'][0]['max_KR_outward_per_corresponding_side_mm']=199
        from policy import digest
        original=p.pop('profile_hash');self.assertNotEqual(digest(p),original)
    def test_equal_200_and_300_allowed(self):self.assertFalse(self.result()['issue'])
    def test_200_on_every_side_is_not_a_sum_limit(self):
        self.e['KR_outward_side_intervals_mm']=[[200,200]]*4;self.assertFalse(self.result()['issue'])
    def test_no_AR_lintels_also_allowed(self):
        self.e['AR_lintel_width_intervals_mm']=[];self.assertFalse(self.result()['issue'])
    def test_KR_continuous_hole_with_AR_lintels(self):
        self.e['AR_lintel_width_intervals_mm']=[[100,100],[300,300]];self.assertFalse(self.result()['issue'])
    def test_KR_200_01_uses_ordinary_RD(self):
        self.e['KR_outward_side_intervals_mm'][1]=[200.01,200.01];self.assertTrue(self.result()['issue'])
    def test_AR_lintel_300_01_uses_ordinary_RD(self):
        self.e['AR_lintel_width_intervals_mm']=[[300.01,300.01]];self.assertTrue(self.result()['issue'])
    def test_uncertainty_at_200_requires_refinement(self):
        self.e['KR_outward_side_intervals_mm'][1]=[199.99,200.01];self.assertIsNone(self.result()['issue'])
    def test_uncertainty_at_300_requires_refinement(self):
        self.e['AR_lintel_width_intervals_mm']=[[299.99,300.01]];self.assertIsNone(self.result()['issue'])
    def test_AR_larger_or_mixed_direction_uses_ordinary_rule(self):
        self.e['AR_main_contour_contained_in_KR_over_whole_depth']=False;self.assertTrue(self.result()['issue'])
    def test_missing_AR_or_KR_never_uses_new_exception(self):
        for side in ['AR','KR']:
            self.e['opening_present_'+side]=False
            self.assertEqual(opening_decision(self.profile,'technical',missing=True,project_geometry_evidence=self.e)['rule'],'absence')
            self.assertTrue(opening_decision(self.profile,'technical',missing=True)['issue'])
    def test_geometry_cannot_be_replaced_by_bounds_only(self):
        self.e.pop('AR_main_contour_contained_in_KR_over_whole_depth');self.assertIsNone(self.result()['issue'])
    def test_missing_whole_opening_evidence_blocks_project_classification(self):
        self.assertIsNone(opening_decision(self.profile,'technical',[0,200,10,30])['issue'])
    def test_unexamined_AR_lintels_do_not_pass(self):
        self.e['all_AR_internal_dividers_examined']=False;self.assertIsNone(self.result()['issue'])
    def test_same_physical_opening_must_be_proved(self):
        self.e['same_physical_opening_proved']=False;self.assertTrue(self.result()['issue'])
    def test_missing_geometry_fingerprint_does_not_pass(self):
        self.e['geometry_evidence_sha256']='';self.assertIsNone(self.result()['issue'])
    def test_negative_width_is_rejected(self):
        self.e['AR_lintel_width_intervals_mm']=[[-1,1]]
        with self.assertRaises(ValueError):self.result()
    def test_reversed_interval_is_rejected(self):
        self.e['KR_outward_side_intervals_mm'][0]=[3,2]
        with self.assertRaises(ValueError):self.result()
    def test_uncertain_outward_direction_is_not_allowed(self):
        self.e['KR_outward_side_intervals_mm'][0]=[-1,1];self.assertIsNone(self.result()['issue'])
    def test_explicit_one_side_absence_cannot_be_lost_without_missing_flag(self):
        for side in ['AR','KR']:
            e=copy.deepcopy(self.e);e['opening_present_'+side]=False
            r=opening_decision(self.profile,'technical',[0,0,0,0],project_geometry_evidence=e)
            self.assertTrue(r['issue']);self.assertEqual(r['rule'],'absence')
    def test_absent_candidate_in_both_disciplines_is_invalid(self):
        self.e['opening_present_AR']=self.e['opening_present_KR']=False
        with self.assertRaises(ValueError):self.result()
    def test_unknown_other_side_is_not_auto_absence(self):
        self.e['opening_present_AR']=False;self.e.pop('opening_present_KR');self.assertIsNone(self.result()['issue'])
    def test_other_project_still_uses_RD_15(self):
        p=resolve({**self.cfg,'project':'sgrvn'});r=opening_decision(p,'technical',[0,16,0,0]);self.assertTrue(r['issue']);self.assertEqual(r['tolerance_mm'],15)
    def test_door_and_window_do_not_get_project_200_even_with_valid_geometry(self):
        for kind in ['door','window']:
            r=opening_decision(self.profile,kind,[0,200,0,0],project_geometry_evidence=self.e)
            self.assertTrue(r['issue']);self.assertEqual(r['rule'],'opposite_pairs');self.assertEqual(r['tolerance_mm'],15)
    def test_door_does_not_need_technical_project_evidence(self):
        r=opening_decision(self.profile,'door',[0,15,0,0]);self.assertFalse(r['issue']);self.assertEqual(r['rule'],'opposite_pairs')
    def test_window_absence_is_still_an_issue(self):self.assertTrue(opening_decision(self.profile,'window',missing=True)['issue'])
    def test_new_rule_is_also_project_scoped_in_PD(self):
        p=resolve({**self.cfg,'stage':'PD','explicit_full_model_test':True});self.assertIn('project_opening_rules',p)
    def test_wall_never_gets_200_even_with_complete_floor_style_evidence(self):
        r=opening_decision(self.profile,'technical',[0,200,10,30],project_geometry_evidence=self.e,host_kind='wall')
        self.assertTrue(r['issue']);self.assertEqual(r['rule'],'four_edges');self.assertEqual(r['tolerance_mm'],15)
        self.assertEqual(r['project_exception']['reason'],'only_floor_openings_get_project_exception')
    def test_wall_equality_15_uses_ordinary_without_project_evidence(self):
        r=opening_decision(self.profile,'technical',[5,5,5,0],host_kind='wall')
        self.assertFalse(r['issue']);self.assertEqual(r['measurement_mm'],15)
    def test_AR_lintels_300_cannot_exempt_a_wall(self):
        self.e['AR_lintel_width_intervals_mm']=[[300,300]]
        r=opening_decision(self.profile,'technical',[0,200,0,0],project_geometry_evidence=self.e,host_kind='wall')
        self.assertTrue(r['issue']);self.assertEqual(r['rule'],'four_edges')
    def test_other_hosts_do_not_get_floor_rule(self):
        r=opening_decision(self.profile,'technical',[0,200,0,0],project_geometry_evidence=self.e,host_kind='other')
        self.assertTrue(r['issue']);self.assertEqual(r['rule'],'four_edges')
    def test_host_cannot_be_inferred_from_a_horizontal_plane(self):
        self.e['normal']=[0,0,1]
        r=opening_decision(self.profile,'technical',[0,200,0,0],project_geometry_evidence=self.e)
        self.assertIsNone(r['issue']);self.assertEqual(r['project_exception']['reason'],'actual_structural_host_kind_required')
    def test_unrecognized_host_kind_is_rejected(self):
        with self.assertRaises(ValueError):opening_decision(self.profile,'technical',[0,200,0,0],project_geometry_evidence=self.e,host_kind='horizontal')
    def test_new_scope_is_part_of_rule_and_profile(self):
        rule=self.profile['project_opening_rules'][0]
        self.assertEqual(rule['host_kinds'],['floor']);self.assertTrue(rule['wall_openings_use_ordinary_stage_and_BEP_rules'])
        self.assertEqual(rule['id'],'sgm_kr_outward_floor_opening_200_ar_lintels_300_v2')
    def test_engine_must_explicitly_implement_new_project_contract(self):
        pf={'ready_for_geometry':True,'profile':self.profile,'mandatory_engine_contract':{},'input_nwd_sha256':'input','fragment_inventory_sha256':'fragments'}
        r={'profile_hash':self.profile['profile_hash'],'input_nwd_sha256':'input','fragment_inventory_sha256':'fragments','complete':True,'implemented_contracts':[]}
        with self.assertRaises(ValueError):validate_engine_receipt(pf,r)
        r['implemented_contracts']=['apply_project_opening_rules_to_matched_whole_openings_and_AR_lintel_surfaces'];validate_engine_receipt(pf,r)
if __name__=='__main__':unittest.main()
