"""Registry identity: project / documentation stage / building / section."""
from pathlib import Path

def scope_key(data):
    result={k:data.get(k) for k in ['project','stage','building_id','section_id']}
    if not result['project'] or result['stage'] not in ['PD','RD']:
        raise ValueError('Explicit project and stage required for registry')
    if not result['building_id'] and not result['section_id']:
        raise ValueError('Explicit building or section required for registry')
    for k,v in result.items():
        if v is None:continue
        if not isinstance(v,str) or not v.strip() or v!=v.strip() or v.endswith('.') or any(c in v for c in '<>:"/\\|?*') or any(ord(c)<32 for c in v) or v in ['.','..','_no_building','_entire_building']:
            raise ValueError('Invalid registry scope component: '+k)
        if v.split('.')[0].upper() in ['CON','PRN','AUX','NUL']+[f'{p}{i}' for p in ['COM','LPT'] for i in range(1,10)]:
            raise ValueError('Reserved registry path component')
    return result

def registry_path(root,data):
    s=scope_key(data)
    return Path(root)/s['project']/s['stage']/(s['building_id'] or '_no_building')/(s['section_id'] or '_entire_building')/'registry.json'
