"""Explicit stage/scope contracts. No geometry inference from drawing names."""
import hashlib,json,math
from pathlib import Path
from registry_scope import scope_key
from result_directory import check_mode
from project_opening_rules import rules_for,evaluate as evaluate_project_opening

PROFILES=json.loads((Path(__file__).resolve().parents[2]/'rules/profiles.json').read_text(encoding='utf8'))

def digest(value):
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()

def resolve(config):
    check_scope=scope_key(config)
    stage=config['stage']
    if stage not in ['PD','RD']:raise ValueError('Stage must be explicit: PD or RD')
    result={**PROFILES[stage],**PROFILES['common'],'stage':stage,'registry_scope':check_scope}
    variant,exclusions=check_mode(config)
    result.update(check_variant=variant,additional_exclusions=exclusions)
    mode=config.get('scope',result['production_scope'])
    if stage=='RD' and mode!='full_model':raise ValueError('RD checks the entire model')
    if stage=='PD' and mode=='full_model' and not config.get('explicit_full_model_test'):
        raise ValueError('Production PD requires supplied plans/sections; full-model test must be explicit')
    if mode not in ['full_model','documented_views']:raise ValueError('Unsupported scope')
    if mode=='documented_views':
        views=config.get('views',[])
        if not views:raise ValueError('Supply and register drawing crops and depths before production PD')
        for view in views:
            if not all(view.get(k) for k in ['id','source','source_sha256','volume','registration_evidence']):
                raise ValueError('View needs source hash, explicit 3D check volume and registration evidence')
            validate_volume(view['volume'])
    result['scope']=mode;result['views']=config.get('views',[]) if mode=='documented_views' else []
    if config.get('fully_constructed'):
        if stage!='RD' or not config.get('building_id') or not config.get('construction_confirmation'):
            raise ValueError('50 mm needs explicit RD building and construction confirmation')
        required=['general_mm','flight_underside_vertical_mm','technical_hole_sum_four_mm','door_window_opposite_pair_mm']
        mapping={k:result['fully_constructed_building_mm'] for k in required}
        if config.get('constructed_tolerance_map',mapping)!=mapping:
            raise ValueError('Legacy mixed tolerance map conflicts with confirmed 50 mm building mode')
        result.update(mapping)
        result['construction_confirmation']=config['construction_confirmation']
        result['fully_constructed']=True
    project_rules=rules_for(config)
    if project_rules:result['project_opening_rules']=project_rules
    result['profile_hash']=digest(result)
    return result

def validate_volume(v):
    # Oriented box in model millimetres. A sheet name or screenshot is not a crop.
    if v.get('type')!='oriented_box':raise ValueError('Registered oriented check volume required')
    origin=v['origin_mm'];basis=v['basis'];bounds=v['bounds_mm']
    if len(origin)!=3 or len(basis)!=3 or any(len(row)!=3 for row in basis) or len(bounds)!=2 or any(len(row)!=3 for row in bounds):raise ValueError('Invalid 3D volume dimensions')
    values=origin+sum(basis,[])+sum(bounds,[])
    if not all(math.isfinite(x) for x in values):raise ValueError('Non-finite crop')
    if any(a>=b for a,b in zip(*bounds)):raise ValueError('Empty check volume')
    for i in range(3):
        for j in range(3):
            if abs(sum(a*b for a,b in zip(basis[i],basis[j]))-(i==j))>1e-8:raise ValueError('Basis must be orthonormal')

def opening_decision(profile,kind,deltas=None,missing=False,*,project_geometry_evidence=None,host_kind=None):
    if missing:return {'issue':True,'rule':'absence','measurement_mm':None,'tolerance_mm':None}
    if kind not in ['technical','door','window']:raise ValueError('Opening type must be classified')
    if kind=='technical' and profile.get('project_opening_rules') and isinstance(project_geometry_evidence,dict):
        presence=[project_geometry_evidence.get('opening_present_'+s) for s in ['AR','KR']]
        if any(p is False for p in presence):
            if all(p is False for p in presence):raise ValueError('No candidate opening exists in either discipline')
            if any(p is True for p in presence):return {'issue':True,'rule':'absence','measurement_mm':None,'tolerance_mm':None}
            return {'issue':None,'rule':'project_geometry_evidence_required','reason':'counterpart_presence_not_proved'}
    if deltas is None or len(deltas)!=4 or not all(math.isfinite(v) for v in deltas):raise ValueError('Four actual boundary deltas required')
    left,right,bottom,top=map(abs,deltas)
    value=left+right+bottom+top if kind=='technical' else max(left+right,bottom+top)
    threshold=profile['technical_hole_sum_four_mm' if kind=='technical' else 'door_window_opposite_pair_mm']
    ordinary={'issue':value>threshold,'measurement_mm':value,'tolerance_mm':threshold,'rule':'four_edges' if kind=='technical' else 'opposite_pairs'}
    if not profile.get('project_opening_rules'):return ordinary
    project=evaluate_project_opening(profile,project_geometry_evidence,kind,host_kind)
    if project['status']=='allowed':return {'issue':False,'rule':project['rule'],'project_exception':project,'ordinary_stage_result':ordinary}
    if project['status']=='pending':return {'issue':None,'rule':'project_geometry_evidence_required','project_exception':project,'ordinary_stage_result':ordinary}
    return {**ordinary,'project_exception':project}

def clip_triangle_to_views(triangle,profile):
    """Exact linear polygon clipping, never classify a whole element by bbox.

    Overlapping view results must be unioned by the geometric engine, retaining
    the original feature identity. Clipping precedes report membership, but the
    *original full solids* supply material/void measurements (no artificial holes).
    """
    if profile['scope']=='full_model':return [{'view_id':'full_model','polygon_mm':triangle}]
    result=[]
    for view in profile['views']:
        v=view['volume'];basis=v['basis'];origin=v['origin_mm'];poly=[list(p) for p in triangle]
        for axis in range(3):
            for bound,sign in [(v['bounds_mm'][0][axis],1),(v['bounds_mm'][1][axis],-1)]:
                def distance(p):return sign*(sum(basis[axis][j]*(p[j]-origin[j]) for j in range(3))-bound)
                output=[]
                for i,a in enumerate(poly):
                    b=poly[(i+1)%len(poly)];da,db=distance(a),distance(b)
                    if da>=0:output.append(a)
                    if (da>=0)!=(db>=0):
                        t=da/(da-db);output.append([a[j]+t*(b[j]-a[j]) for j in range(3)])
                poly=output
                if not poly:break
            if not poly:break
        if len(poly)>=3:
            a=poly[0];area=0
            for i in range(1,len(poly)-1):
                u=[poly[i][j]-a[j] for j in range(3)];w=[poly[i+1][j]-a[j] for j in range(3)]
                area+=sum(x*x for x in [u[1]*w[2]-u[2]*w[1],u[2]*w[0]-u[0]*w[2],u[0]*w[1]-u[1]*w[0]])**.5/2
            if area>1e-12:result.append({'view_id':view['id'],'polygon_mm':poly})
    return result
