"""Scoped project opening policy applied to already-certified geometry evidence.

This module does not derive a contour, identify a lintel, or prove containment
from bounding boxes. The geometry adapter must validate the evidence against
the current NWD and retain all referenced proof artifacts in its receipt.
"""
import json,math
from pathlib import Path

RULE_PATH=Path(__file__).resolve().parents[2]/'rules/projects/sochi_moskva.json'
RULE=json.loads(RULE_PATH.read_text(encoding='utf8'))

def rules_for(config):
    if config.get('project')==RULE['project'] and config.get('stage') in RULE['stages']:
        return [json.loads(json.dumps(RULE,ensure_ascii=False))]
    return []

def interval(value):
    if not isinstance(value,(list,tuple)) or len(value)!=2:raise ValueError('A certified lower/upper interval is required')
    if any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) for x in value):raise ValueError('Interval must contain finite numbers')
    a,b=value
    if a>b:raise ValueError('Interval order differs')
    return a,b

def evaluate(profile,evidence,kind='technical',host_kind=None):
    active=profile.get('project_opening_rules',[])
    if not active:return {'status':'not_applicable','reason':'no_project_rule'}
    if len(active)!=1 or active[0].get('id')!=RULE['id'] or profile['registry_scope']['project']!=RULE['project']:
        raise ValueError('Unrecognized or unscoped project rule')
    if kind not in active[0]['opening_kinds']:return {'status':'not_applicable','reason':'door_window_stage_rules_unchanged','rule':RULE['id']}
    if host_kind is None:return {'status':'pending','reason':'actual_structural_host_kind_required','rule':RULE['id']}
    if host_kind not in ['floor','wall','other']:raise ValueError('Host kind must be floor, wall or other')
    if host_kind not in active[0]['host_kinds']:
        return {'status':'not_applicable','reason':'only_floor_openings_get_project_exception','rule':RULE['id']}
    if not isinstance(evidence,dict):return {'status':'pending','reason':'certified_whole_opening_evidence_required','rule':RULE['id']}
    for name in ['opening_present_AR','opening_present_KR']:
        if evidence.get(name) is False:return {'status':'not_applicable','reason':'missing_opening_uses_ordinary_rules','rule':RULE['id']}
        if evidence.get(name) is not True:return {'status':'pending','reason':'opening_presence_not_proved','rule':RULE['id']}
    for name in ['same_physical_opening_proved','AR_main_contour_contained_in_KR_over_whole_depth']:
        if evidence.get(name) is False:return {'status':'not_applicable','reason':name+'_not_satisfied','rule':RULE['id']}
        if evidence.get(name) is not True:return {'status':'pending','reason':name+'_not_proved','rule':RULE['id']}
    sha=evidence.get('geometry_evidence_sha256','')
    if not isinstance(sha,str) or len(sha)!=64 or any(x not in '0123456789abcdef' for x in sha):
        return {'status':'pending','reason':'geometry_proof_artifact_fingerprint_required','rule':RULE['id']}
    if evidence.get('all_AR_internal_dividers_examined') is not True:
        return {'status':'pending','reason':'AR_lintels_not_completely_examined','rule':RULE['id']}
    offsets=evidence.get('KR_outward_side_intervals_mm');widths=evidence.get('AR_lintel_width_intervals_mm')
    if not isinstance(offsets,list) or len(offsets)!=4 or not isinstance(widths,list):
        return {'status':'pending','reason':'all_four_side_intervals_and_all_lintel_widths_required','rule':RULE['id']}
    offsets=[interval(x) for x in offsets];widths=[interval(x) for x in widths]
    if any(b<0 for a,b in offsets) or any(a<0 for a,b in widths):raise ValueError('Evidence contradicts nonnegative containment/width')
    if any(a<0 for a,b in offsets):
        return {'status':'pending','reason':'outward_direction_not_strictly_certified','rule':RULE['id']}
    cap=active[0]['max_KR_outward_per_corresponding_side_mm'];bridge=active[0]['max_disregarded_AR_lintel_width_mm']
    if any(a>cap for a,b in offsets):return {'status':'not_applicable','reason':'KR_outward_side_exceeds_project_limit','rule':RULE['id']}
    if any(a>bridge for a,b in widths):return {'status':'not_applicable','reason':'AR_lintel_exceeds_project_limit','rule':RULE['id']}
    if any(b>cap for a,b in offsets) or any(b>bridge for a,b in widths):
        return {'status':'pending','reason':'uncertainty_crosses_project_limit','rule':RULE['id']}
    return {'status':'allowed','rule':RULE['id'],'max_KR_outward_per_side_mm':cap,
            'max_disregarded_AR_lintel_width_mm':bridge,'geometry_evidence_sha256':sha,
            'applies_only_to_this_matched_opening':True,'source_geometry_not_modified':True}
