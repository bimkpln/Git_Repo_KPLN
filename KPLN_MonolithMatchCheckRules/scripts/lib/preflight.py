"""Mandatory admission contract before a new NWD geometry run."""
import argparse,json
from pathlib import Path
from policy import resolve,digest
from materials import select_fragments
from history import atomic_write

def raw_fragments(raw):
    if any('fragmentRecords' in item for item in raw):
        from nwd_import import fragments_from_nodes
        return fragments_from_nodes(raw)
    if any(item.get('fragments',1)!=1 for item in raw):
        raise ValueError('Legacy multi-fragment node export cannot prove fragment ownership/materials')
    records=[]
    for item in raw:
        ancestry=item.get('ancestry',[]);props={}
        # Nearest element properties win over system/type ancestors.
        for node in reversed(ancestry):
            for p in node.get('properties',[]):
                if p.get('category')=='Объект':props[p['name']]=str(p.get('value','')).partition(':')[2]
        material=''
        for node in ancestry:
            matches=[p for p in node.get('properties',[]) if (p.get('category')=='Материал' and p.get('name')=='Имя') or p.get('name')=='Материал']
            if matches:material=str(matches[0].get('value','')).partition(':')[2];break
        names=' '.join(str(props.get(k,'')) for k in ['Семейство','Тип','Имя'])
        # Do not use visibility alone: incidental hidden view state is not the
        # evaluated recursive membership of the requested search-set folder.
        method=props.get('Способ изготовления') or props.get('Способ возведения') or props.get('Construction Method','')
        if not method and ('монолит' in names.lower() or 'сборн' in names.lower()):method=names
        records.append({'fragment_id':str(item['index']),'element_id':props.get('Id'),'category':props.get('Категория',''),
            'material':material,'construction_method':method,'excluded_by_hide_set':bool(item.get('excludedBySet')),
            'is_void_cutter':props.get('Тип геометрии','').lower() in ['вырез','void'],
            'role':props.get('MonolithCheckRole',''),'triangle_offset':item.get('offset'),'triangle_count':item.get('triangles')})
    return records

def prepare(config,fragments):
    profile=resolve(config)
    fragments=[dict(f) for f in fragments]
    for assignment in config.get('verified_fragment_assignments',[]):
        target=next((f for f in fragments if f['fragment_id']==assignment['fragment_id']),None)
        evidence=assignment.get('evidence',{})
        values=assignment.get('values',{})
        if (target is None or assignment.get('input_nwd_sha256')!=config['input_nwd_sha256']
            or assignment.get('fragment_metadata_sha256')!=digest(target)
            or evidence.get('registry_scope')!=profile['registry_scope']
            or evidence.get('type')!='user_confirmation' or not evidence.get('statement')
            or not values or not set(values)<={'material','construction_method','role'}
            or not all(isinstance(v,str) and v.strip() for v in values.values())):
            raise ValueError('Stale, unscoped or unsupported fragment assignment')
        target.update(values,assignment_evidence=evidence)
    for certificate in config.get('verified_original_closures',[]):
        target=next((f for f in fragments if f['fragment_id']==certificate['fragment_id']),None)
        if target is None or certificate.get('input_nwd_sha256')!=config['input_nwd_sha256'] or certificate.get('fragment_metadata_sha256')!=digest(target) or not certificate.get('proof_sha256'):
            raise ValueError('Stale closure/material certificate')
        target['verified_original_closure']=certificate['proof_sha256']
    materials=config.get('verified_monolith_materials',[])
    if materials and not config.get('material_dictionary_evidence'):raise ValueError('Project material dictionary requires recorded evidence')
    selection=select_fragments(fragments,materials)
    result={'schema':2,**profile['registry_scope'],'profile':profile,
        'input_nwd_sha256':config['input_nwd_sha256'],'fragment_inventory_sha256':digest(fragments),
        'material_selection':selection,'ready_for_geometry':selection['complete'],
        'mandatory_engine_contract':{
          'apply_explicit_additional_exclusions':bool(profile['additional_exclusions']),
          'selected_fragments_only_before_boolean_union':True,
          'unknown_materials_must_not_become_absent_counterparts':True,
          'measure_against_original_full_solids_before_scope_clipping':True,
          'clip_discrepancy_geometry_to_registered_view_volumes':profile['scope']=='documented_views',
          'union_overlapping_views_without_duplicate_issues':True,
          'profile_hash_in_geometry_and_cache_receipts':profile['profile_hash'],
          'single_history_reconciled_hierarchy_for_all_reports':True,
          'release_blocked_until_engine_receipt_matches_contract':True}}
    if profile.get('project_opening_rules'):
        result['mandatory_engine_contract']['apply_project_opening_rules_to_matched_whole_openings_and_AR_lintel_surfaces']=True
    return result

def validate_engine_receipt(preflight,receipt):
    if not preflight['ready_for_geometry']:raise ValueError('Unclassified model fragments remain')
    expected={k for k,v in preflight['mandatory_engine_contract'].items() if v is True}
    if preflight['profile'].get('project_opening_rules'):expected.add('apply_project_opening_rules_to_matched_whole_openings_and_AR_lintel_surfaces')
    if receipt.get('profile_hash')!=preflight['profile']['profile_hash'] or receipt.get('input_nwd_sha256')!=preflight['input_nwd_sha256'] or receipt.get('fragment_inventory_sha256')!=preflight['fragment_inventory_sha256']:
        raise ValueError('Geometry engine ran against another profile/input')
    if not expected<=set(receipt.get('implemented_contracts',[])):raise ValueError('Engine lacks required stage/filter/scope/history integration')
    if not receipt.get('complete'):raise ValueError('Incomplete geometry result')

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('config');p.add_argument('raw_nwd_json');p.add_argument('output');a=p.parse_args()
    config=json.loads(Path(a.config).read_text(encoding='utf8'));raw=json.loads(Path(a.raw_nwd_json).read_text(encoding='utf8'))
    result=prepare(config,raw_fragments(raw));atomic_write(a.output,result)
    print(json.dumps({'ready_for_geometry':result['ready_for_geometry'],'counts':result['material_selection']['counts']},ensure_ascii=False))
    if not result['ready_for_geometry']:raise SystemExit(2)
