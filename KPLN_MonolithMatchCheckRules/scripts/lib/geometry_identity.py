"""Conservative identity from certified discrepancy geometry, never a bbox.

Equivalent geometry with different serialization may require an explicit match
certificate. This false ambiguity is preferable to an incorrect automatic match.
NPZ archive timestamps, file names and image paths do not affect identity.
"""
import copy,hashlib,zipfile
from pathlib import Path
from policy import digest

def bind_geometry(entries,run_directory):
    result=copy.deepcopy(entries);root=Path(run_directory).resolve();cache={}
    for entry in result:
        for g in entry.get('local_instances') or [entry]:
            sections=g.get('certified_sections',[]);meshes={}
            for side,paths in g.get('certified_meshes',{}).items():
                hashes=[]
                for relative in paths:
                    p=(root/relative).resolve()
                    if not p.is_relative_to(root):raise ValueError('Geometry proof outside run directory')
                    if p not in cache:
                        with zipfile.ZipFile(p) as archive:
                            names=sorted(archive.namelist())
                            if not names:raise ValueError('Empty geometry proof')
                            cache[p]=digest({n:hashlib.sha256(archive.read(n)).hexdigest() for n in names})
                    hashes.append(cache[p])
                if hashes:meshes[side]=sorted(hashes)
            if not sections and not meshes:raise ValueError('No certified geometry for '+g['number'])
            g['identity_geometry_sha256']=digest({'sections':sections,'meshes':meshes,
                'association_regions':g.get('association_regions'),
                'basis':g.get('basis'),'offset':g.get('offset'),'opening_presence':g.get('opening_presence')})
    return result
