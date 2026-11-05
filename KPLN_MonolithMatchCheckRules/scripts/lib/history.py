"""Immutable observations and stable identifiers shared by both report types.

    Exact unchanged features match automatically. Changed or ambiguous shapes
    require a geometry-backed match certificate, never nearest-ID guessing.
"""
import copy,datetime,hashlib,json,os
from pathlib import Path
from policy import digest
from registry_scope import scope_key

def owners(g):
    return sorted({(s,c[s]) for c in g['children'] for s in ['AR','KR'] if c.get(s) and c.get('missing_side')!=s})

def feature(g):
    proof=g.get('identity_geometry_sha256','')
    if len(proof)!=64 or any(c not in '0123456789abcdef' for c in proof):
        raise ValueError('History identity requires a fingerprint of certified geometry, not only IDs/bounds')
    return digest({k:g.get(k) for k in ['category','measurement_rule','bounds_mm','source','opening_presence','association_shape','association_depth','association_basis','identity_geometry_sha256']}|{'owners':owners(g)})

def observed(g,run_id,date):
    return {'run_id':run_id,'date':date,'feature':feature(g),'bounds_mm':g['bounds_mm'],
            'owners':owners(g),'category':g['category'],'measurement_rule':g['measurement_rule'],
            'source_number':g['number'],'measurement_mm':g.get('measurement_mm')}

def atomic_write(path,data):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    temp=path.with_suffix(path.suffix+'.tmp');temp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8');os.replace(temp,path)

def seed(entries,project,stage,run_id,date,source_sha256,*,building_id=None,section_id=None):
    datetime.date.fromisoformat(date)
    scope=scope_key(dict(project=project,stage=stage,building_id=building_id,section_id=section_id))
    reg={'schema':2,**scope,'source_sha256':source_sha256,
         'runs':[{'id':run_id,'date':date,'kind':'verified_baseline','earlier_history_complete':False}],
         'issues':{},'groups':{}}
    for entry in entries:
        members=[]
        for g in entry.get('local_instances') or [entry]:
            uid=digest([scope,run_id,g['number']])[:24];members.append(uid)
            if uid in reg['issues']:raise ValueError('Duplicate physical issue')
            reg['issues'][uid]={'number':g['number'],'first_detected':date,'first_detected_basis':'first_verified_registry_baseline; earlier_history_not_reconstructed',
                'status':'baseline','last_confirmed_status':'baseline',
                'status_events':[{'run_id':run_id,'date':date,'status':'baseline'}],
                'observations':[observed(g,run_id,date)],
                'markers':copy.deepcopy(g['children'])}
        reg['groups'][entry['number']]={'members':members,'first_detected':date,'active':True}
    return reg

def close_regions(a,b):
    return all(a[1][i]+.5>=b[0][i] and b[1][i]+.5>=a[0][i] for i in range(3))

def reconcile(registry,physical,run_id,date,coverage,certificates=()):
    """coverage is emitted by successful geometry/scope validation, not the UI.

    No finding may be marked resolved under changed rules, unknown materials,
    missing source section, partial feature coverage, or failed computation.
    Idempotent run IDs are not new observations.
    """
    datetime.date.fromisoformat(date)
    if registry.get('superseded_by'):raise ValueError('Registry relocated; use '+registry['superseded_by'])
    if any(r['id']==run_id for r in registry['runs']):raise ValueError('Run already registered')
    if any(date<r['date'] for r in registry['runs']):raise ValueError('Observation date precedes history')
    if scope_key(coverage)!=scope_key(registry):raise ValueError('Separate registry required for another project/stage/building/section')
    if not coverage.get('complete') or not coverage.get('material_filter_complete') or not coverage.get('geometry_receipt_sha256'):
        raise ValueError('Incomplete geometry/material classification cannot update issue history')
    old=registry['issues'];matches={};used=set();ambiguous=[]
    bynumber={g['number']:g for g in physical}
    if len(bynumber)!=len(physical):raise ValueError('Duplicate physical numbers')
    for g in physical:
        exact=[uid for uid,v in old.items() if v['observations'][-1]['feature']==feature(g)]
        if len(exact)==1 and exact[0] not in used:matches[g['number']]=exact[0];used.add(exact[0])
        elif exact:ambiguous.append({'current':g['number'],'candidates':exact,'reason':'non_unique_exact_match'})
    for c in certificates:
        if c.get('relation')!='same_physical_feature' or not c.get('evidence_sha256'):raise ValueError('A changed shape needs a geometry-backed match certificate')
        uid=c['previous_uid'];number=c['current_number']
        if uid not in old or number not in bynumber or number in matches or uid in used:raise ValueError('Non-bijective match')
        if c.get('previous_feature')!=old[uid]['observations'][-1]['feature'] or c.get('current_feature')!=feature(bynumber[number]):raise ValueError('Stale matching certificate')
        matches[number]=uid;used.add(uid)
    for g in physical:
        if g['number'] in matches:continue
        possible=[uid for uid,v in old.items() if uid not in used and v['observations'][-1]['category']==g['category'] and
                  set(map(tuple,v['observations'][-1]['owners'])) & set(owners(g)) and close_regions(v['observations'][-1]['bounds_mm'],g['bounds_mm'])]
        if possible:ambiguous.append({'current':g['number'],'candidates':possible,'reason':'changed_feature_needs_geometry_match'})
    if ambiguous:return {'complete':False,'ambiguous':ambiguous,'registry':None}
    result=copy.deepcopy(registry);mapping={};reserved={v['number'] for v in old.values()}|set(registry['groups'])
    def next_number():
        prefix=coverage['number_prefix'];i=1
        while f'{prefix}{i:03d}' in reserved:i+=1
        value=f'{prefix}{i:03d}';reserved.add(value);return value
    for g in physical:
        uid=matches.get(g['number'])
        if uid is None:
            uid=digest([scope_key(registry),run_id,g['number'],feature(g)])[:24]
            stable=next_number();markers=copy.deepcopy(g['children'])
            for i,c in enumerate(markers,1):c['number']=stable+(f'.{i:02d}' if len(markers)>1 else '')
            result['issues'][uid]={'number':stable,'first_detected':date,'first_detected_basis':'complete_verified_run','observations':[],'markers':markers}
            status='new'
        else:status='reopened' if old[uid].get('last_confirmed_status',old[uid]['status'])=='resolved' else 'unworked'
        item=result['issues'][uid];item['status']=status;item['last_confirmed_status']=status
        item.setdefault('status_events',[]).append({'run_id':run_id,'date':date,'status':status})
        item['observations'].append(observed(g,run_id,date));mapping[g['number']]=uid
    for uid,item in result['issues'].items():
        if uid in mapping.values():continue
        # Old issue IDs must be explicitly covered in their entirety; a crop
        # intersection or disappeared Revit element is not proof of resolution.
        covered=coverage.get('same_rules') is True and uid in coverage.get('fully_rechecked_issue_uids',[])
        item['status']='resolved' if covered else 'not_checked'
        if covered:item['last_confirmed_status']='resolved'
        item.setdefault('status_events',[]).append({'run_id':run_id,'date':date,'status':item['status'],
            'reason':'fully_rechecked_same_rules' if covered else 'no_complete_same_rule_coverage'})
    result['runs'].append({'id':run_id,'date':date,'coverage':coverage})
    return {'complete':True,'ambiguous':[],'registry':result,'physical_uid_map':mapping}

def export_history(entries,registry,uid_map):
    """One canonical hierarchy for plugin and coordinator; never regroup in UI.

    A changed group reuses its old number only when the overlap is unambiguous.
    Splits/merges return a review condition instead of silently renumbering.
    """
    result=copy.deepcopy(entries);oldgroups=registry['groups'];claimed=set();groups_out={}
    for e in result:
        local=e.get('local_instances') or [e];members=[uid_map[g['number']] for g in local]
        candidates=[n for n,v in oldgroups.items() if set(v['members'])&set(members)]
        if len(candidates)>1 or any(n in claimed for n in candidates):raise ValueError('Split/merged historical group needs explicit lineage decision')
        parent=candidates[0] if candidates else registry['issues'][members[0]]['number'];claimed.add(parent)
        groups_out[parent]={'members':members,'active':True,'first_detected':min(registry['issues'][u]['first_detected'] for u in members)}
        children=[];seen_markers=set()
        for g,uid in zip(local,members):
            history=registry['issues'][uid];g['number']=history['number'];g['parent_number']=parent
            g['first_detected']=history['first_detected'];g['lifecycle_status']=history['status'];g['issue_uid']=uid
            prior=history['markers']
            for c in g['children']:
                pair=lambda x:(x.get('AR'),x.get('KR'),x.get('missing_side'))
                candidates=[x for x in prior if pair(x)==pair(c)]
                if len(candidates)>1:
                    candidates=[x for x in candidates if x.get('point_mm')==c.get('point_mm')]
                if len(candidates)!=1:raise ValueError('Marker split/merge needs explicit lineage; never guess or reuse a number')
                c['number']=candidates[0]['number']
                if c['number'] in seen_markers:raise ValueError('Duplicate persistent marker number')
                seen_markers.add(c['number']);c['first_detected']=g['first_detected'];c['lifecycle_status']=g['lifecycle_status'];children.append(c)
        e['number']=parent;e['children']=children;e['local_instances']=local if local[0] is not e else [copy.deepcopy({k:v for k,v in e.items() if k!='local_instances'})]
        states={registry['issues'][u]['status'] for u in members}
        e['first_detected']=groups_out[parent]['first_detected'];e['lifecycle_status']='new' if 'new' in states else 'baseline' if states=={'baseline'} else 'reopened' if 'reopened' in states else 'unworked'
        e['entry_type']='group' if len(children)>1 else 'single'
    return result,groups_out
