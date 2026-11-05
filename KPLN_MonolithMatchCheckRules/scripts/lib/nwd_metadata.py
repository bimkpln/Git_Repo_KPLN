"""Validate metadata structure and source identity, never geometric readiness."""
import json,math
from pathlib import Path
from nwd_import import filehash
from stream_sequence import items

def validate_metadata_export(directory,expected_input_sha256,expected_reader_sha256):
    directory=Path(directory);manifest_path=directory/'manifest.json';nodes_path=directory/'nodes.json'
    m=json.loads(manifest_path.read_text(encoding='utf8'))
    if m.get('schema')!=1 or m.get('complete') is not True:raise ValueError('Complete metadata schema 1 required')
    if Path(m.get('input','')).suffix.casefold()!='.nwd':raise ValueError('Only NWD metadata is admitted')
    if m.get('input_sha256')!=expected_input_sha256 or m.get('input_sha256_after_read')!=expected_input_sha256:raise ValueError('Input NWD changed or differs')
    if m.get('reader_executable_sha256')!=expected_reader_sha256:raise ValueError('Unverified metadata reader executable')
    if m.get('metadata_sha256')!=filehash(nodes_path):raise ValueError('Metadata bytes differ')
    if m.get('geometry_extracted') is not False or m.get('source_modified') is not False:raise ValueError('Metadata-only read required')
    if not isinstance(m.get('models'),list) or any(not isinstance(r.get('file'),str) or not isinstance(r.get('units'),str) or not r['units'] for r in m['models']):raise ValueError('Explicit model names and units required')
    model_names={r['file'] for r in m['models']}
    count=0;excluded=0;sources=[]
    for r in items(nodes_path):
        if not isinstance(r,dict) or any(not isinstance(r.get(k),str) for k in ['name','model','source']):raise ValueError('Metadata names and source strings required')
        if r['model'] not in model_names:raise ValueError('Unknown metadata model')
        if type(r.get('index')) is not int or r['index']!=count:raise ValueError('Metadata node sequence differs')
        parent=r.get('parent')
        if type(parent) is not int or parent< -1 or parent>=count:raise ValueError('Metadata parent sequence differs')
        if not isinstance(r.get('properties'),list):raise ValueError('Metadata properties array required')
        if type(r.get('has_geometry')) is not bool or type(r.get('hidden')) is not bool:raise ValueError('Explicit geometry and visibility flags required')
        expected_source=r['name'] if r['name'].casefold().endswith('.nwc') else sources[parent] if parent>=0 else r['model']
        if r.get('source')!=expected_source:raise ValueError('Metadata source hierarchy differs')
        sources.append(expected_source)
        bounds=r.get('bounds')
        if bounds is not None:
            if len(bounds)!=6 or not all(isinstance(v,(int,float)) and math.isfinite(v) for v in bounds):raise ValueError('Nonfinite metadata bounds')
            if any(bounds[i]>bounds[i+3] for i in range(3)):raise ValueError('Reversed metadata bounds')
        if type(r.get('excluded_by_hide_set')) is not bool:raise ValueError('Explicit evaluated Hide-set membership required')
        excluded+=int(r['excluded_by_hide_set']);count+=1
    if type(m.get('nodes')) is not int or count!=m['nodes']:raise ValueError('Metadata node count differs')
    if excluded!=m.get('excluded_set_nodes'):raise ValueError('Metadata Hide-set coverage differs')
    return {'metadata_import_complete':True,'scene_nodes':count,'excluded_set_nodes':excluded,
        'input_nwd_sha256':expected_input_sha256,'reader_executable_sha256':expected_reader_sha256,
        'metadata_sha256':m['metadata_sha256'],'manifest_sha256':filehash(manifest_path),
        'geometry_check_performed':False,'semantic_correspondence_proved':False,'complete_check':False}
