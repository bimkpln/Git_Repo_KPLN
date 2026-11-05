"""Validate an actual per-fragment NWD export before any material union.

This is an admission check, never a geometric comparison or release receipt.
Legacy node-aggregated exports cannot prove per-fragment material ownership.
"""
import hashlib,json,math
from pathlib import Path
from policy import digest
from json_array_stream import array_items
from nwd_literal_matrices import LiteralMatrices

def filehash(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda:f.read(4*1024*1024),b''):h.update(block)
    return h.hexdigest()

def value(v):
    text=str(v)
    return text.partition(':')[2] if ':' in text else text

def owner_properties(ancestry):
    # Select the nearest INSTANCE with an ID. Do not combine its identity with
    # properties of another instance/system higher in the ancestry.
    for node in ancestry:
        props={p['name']:value(p.get('value','')) for p in node.get('properties',[]) if p.get('category')=='Объект'}
        if props.get('Id'):return props
    return {}

def fragments_from_nodes(nodes):
    result=[]
    for node in nodes:
        ancestry=node.get('ancestry',[]);props=owner_properties(ancestry)
        source=next((n['name'] for n in ancestry if n.get('name','').lower().endswith('.nwc')),None)
        fragments=node.get('fragmentRecords')
        if not isinstance(fragments,list) or len(fragments)!=node.get('fragments'):
            raise ValueError('Missing individual fragment records for node '+str(node.get('index')))
        if sum(f.get('triangles',-1) for f in fragments)!=node.get('triangles'):
            raise ValueError('Fragment triangle coverage mismatch')
        cursor=node['offset']
        for f in fragments:
            matrix=f.get('local_to_world',[])
            if len(matrix)!=16 or not all(isinstance(x,(int,float)) and math.isfinite(x) for x in matrix):
                raise ValueError('Invalid fragment transformation')
            if any(abs(matrix[k])>1e-12 for k in [3,7,11]) or abs(matrix[15]-1)>1e-12:
                raise ValueError('Non-affine fragment transformation')
            determinant=(matrix[0]*(matrix[5]*matrix[10]-matrix[9]*matrix[6])
                -matrix[4]*(matrix[1]*matrix[10]-matrix[9]*matrix[2])
                +matrix[8]*(matrix[1]*matrix[6]-matrix[5]*matrix[2]))
            if determinant==0:raise ValueError('Singular fragment transformation')
            if not f.get('path_matches_node'):raise ValueError('Fragment owner path differs from exported node')
            if f.get('excludedBySet')!=node.get('excludedBySet'):
                raise ValueError('Fragment/node hide membership differs')
            if f.get('offset')!=cursor or not isinstance(f.get('triangles'),int) or f['triangles']<0:
                raise ValueError('Non-contiguous fragment byte range')
            cursor+=f['triangles']*72
            material=f.get('material','')
            evidence=f.get('material_evidence',[])
            if material and not evidence:raise ValueError('Material has no source evidence')
            if any(e.get('value')!=material for e in evidence):raise ValueError('Conflicting fragment material evidence')
            names=' '.join(str(props.get(k,'')) for k in ['Семейство','Тип','Имя'])
            method=props.get('Способ изготовления') or props.get('Способ возведения') or props.get('Construction Method','')
            if not method and ('монолит' in names.lower() or 'сборн' in names.lower()):method=names
            result.append({'fragment_id':f['fragment_id'],'node_index':node['index'],
                'element_id':props.get('Id'),'source':source,'category':props.get('Категория',''),
                'material':material,'material_evidence':evidence,'construction_method':method,
                'excluded_by_hide_set':f['excludedBySet'],'triangle_offset':f['offset'],
                'triangle_count':f['triangles'],'local_to_world':matrix,
                'is_void_cutter':props.get('Тип геометрии','').lower() in ['вырез','void'],
                'role':props.get('MonolithCheckRole',''),'owner_properties':props})
    ids=[f['fragment_id'] for f in result]
    if len(ids)!=len(set(ids)):raise ValueError('Duplicate fragment identity')
    return result

def validate_export(directory,input_sha256):
    directory=Path(directory)
    paths={k:directory/n for k,n in {'manifest':'nwd_export_manifest.json','nodes':'nwd_geometry.json',
        'triangles':'nwd_triangles.bin','sets':'nwd_sets.json','scene':'all_nodes.json'}.items()}
    manifest=json.loads(paths['manifest'].read_text(encoding='utf-8-sig'))
    if manifest.get('complete') is not True or manifest.get('schema')!=2:
        raise ValueError('Complete fragment export schema 2 required')
    if manifest.get('input_sha256')!=input_sha256:raise ValueError('Export belongs to another NWD')
    cursor=0;node_count=0;hidden_nodes=0
    matrices=LiteralMatrices(directory,manifest)
    def checked_nodes():
        nonlocal cursor,node_count,hidden_nodes
        for node in array_items(paths['nodes']):
            if node['index']!=node_count:raise ValueError('Non-contiguous node identity')
            if node['offset']!=cursor:raise ValueError('Node byte ranges overlap or leave gaps')
            cursor+=node['triangles']*72;node_count+=1;hidden_nodes+=bool(node['excludedBySet'])
            yield matrices.node(node)
    try:
        fragments=fragments_from_nodes(checked_nodes())
        matrix_reference=matrices.finish()
    finally:matrices.close()
    if manifest.get('geometry_nodes')!=node_count:raise ValueError('Incomplete geometry node inventory')
    if cursor!=paths['triangles'].stat().st_size:raise ValueError('Truncated or trailing triangle bytes')
    if matrix_reference is not None:paths['literal_matrices']=directory/matrix_reference['file']
    from nwd_local_coordinates import validate_local_reference
    local_reference=validate_local_reference(directory,manifest,fragments)
    if local_reference is not None:paths['local_triangles']=directory/local_reference['file']
    scene_count=0;scene_geometry=0
    for node in array_items(paths['scene'],wrapper_key='nodes'):
        scene_count+=1;scene_geometry+=bool(node['has_geometry'])
    if scene_count!=manifest.get('scene_nodes'):raise ValueError('Incomplete full scene tree')
    if scene_geometry!=node_count:raise ValueError('Scene geometry inventory mismatch')
    sets=json.loads(paths['sets'].read_text(encoding='utf-8-sig'))
    # Missing sources/IDs on real material block admission, including hidden
    # geometry in inventory. Zero-triangle annotation nodes remain recorded.
    unknown=[f['fragment_id'] for f in fragments if f['triangle_count'] and not f['excluded_by_hide_set'] and (not f['source'] or not f['element_id'])]
    receipt={'schema':1,'import_complete':not unknown,'input_nwd_sha256':input_sha256,
        'geometry_nodes':node_count,'scene_nodes':scene_count,'fragments':len(fragments),
        'triangles':cursor//72,'hidden_geometry_nodes':hidden_nodes,
        'hidden_fragments':sum(f['excluded_by_hide_set'] for f in fragments),
        'hide_search_sets':[s for s in sets if s.get('excluded')],
        'sources':sorted({f['source'] for f in fragments if f['source']}),
        'unresolved_source_or_instance':unknown,'fragment_inventory_sha256':digest(fragments),
        'artifact_hashes':{str(p.name):filehash(p) for p in paths.values()},
        'scope':'export structure, identity, materials and byte coverage; no solid geometry certification',
        'complete_geometry_check':False,'release_allowed':False}
    if local_reference is not None:receipt['local_coordinates_reference']=local_reference
    if matrix_reference is not None:receipt['literal_matrices_reference']=matrix_reference
    return fragments,receipt
