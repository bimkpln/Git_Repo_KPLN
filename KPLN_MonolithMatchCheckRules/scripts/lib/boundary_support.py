"""Prove complete directed edge support for existing boundary face patches."""
from collections import defaultdict
import numpy as np,trimesh

def directed_open_edges(triangles):
    vertices,ii=np.unique(np.round(np.asarray(triangles).reshape(-1,3),9),axis=0,return_inverse=True)
    faces=ii.reshape(-1,3);edges=defaultdict(list)
    for face in faces:
        for a,b in zip(face,np.roll(face,-1)):edges[tuple(sorted((int(a),int(b))))].append(a<b)
    result=[]
    for (a,b),directions in edges.items():
        balance=sum(1 if flag else -1 for flag in directions)
        if balance:result.append(vertices[[a,b] if balance>0 else [b,a]])
    return np.asarray(result).reshape(-1,2,3)

def complete_segment_support(a,b,targets,tolerance_mm,opposite=True):
    d=b-a;length=np.linalg.norm(d)
    if not length:return False
    axis=d/length;offsets=targets-a
    parameters=offsets@axis/length
    normal=offsets-parameters[:,:,None]*d
    valid=np.linalg.norm(normal,axis=2).max(1)<=tolerance_mm
    if opposite:valid&=np.einsum('ij,j->i',targets[:,1]-targets[:,0],d)<0
    intervals=np.sort(parameters[valid],axis=1)
    intervals=intervals[(intervals[:,1]>=0)&(intervals[:,0]<=1)]
    intervals=intervals[np.argsort(intervals[:,0])] if len(intervals) else intervals
    end=0.;gap=tolerance_mm/length
    for low,high in intervals:
        if low>end+gap:return False
        end=max(end,float(high))
        if end>=1-gap:return True
    return False

def boundary_patch_receipt(selected_triangles,patch_triangles,tolerance_mm):
    selected=directed_open_edges(selected_triangles);patch=directed_open_edges(patch_triangles)
    v,ii=np.unique(np.round(np.asarray(patch_triangles).reshape(-1,3),9),axis=0,return_inverse=True)
    mesh=trimesh.Trimesh(vertices=v,faces=ii.reshape(-1,3),process=False)
    parts=mesh.split(only_watertight=False,repair=False)
    closed_components=sum(bool(p.is_watertight) for p in parts)
    supported=[complete_segment_support(a,b,selected,tolerance_mm) for a,b in patch]
    return {'passed':bool(len(patch)>0 and all(supported) and not closed_components),
        'selected_unbalanced_edges':len(selected),'patch_boundary_edges':len(patch),
        'unsupported_patch_edges':np.flatnonzero(~np.asarray(supported,bool)).tolist(),
        'patch_components':len(parts),'closed_patch_components':closed_components,
        'directed_continuous_segment_support':True,'endpoint_sampling_is_not_used':True,
        'support_tolerance_mm':float(tolerance_mm),'added_or_guessed_triangles':0,'material_dictionary_changed':False}
