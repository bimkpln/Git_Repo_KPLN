"""Clipping faces are identified by input ownership, then checked spatially."""
import numpy as np

def classify_faces(triangles,original_ids,source_owners,clip_ids,half,guard_mm):
    tri=np.asarray(triangles,float);ids=np.asarray(original_ids,np.int64);clip_ids=set(clip_ids)
    if set(source_owners)&clip_ids:raise ValueError('Source and clipping input identities overlap')
    if guard_mm<0 or half<=0 or len(tri)!=len(ids):raise ValueError('Invalid clipping proof inputs')
    owners=np.array([source_owners.get(int(i),-1) for i in ids],dtype=np.int32)
    caps=np.isin(ids,list(clip_ids));unknown=(owners<0)&~caps
    if unknown.any():raise ValueError('Face has no source or clipping input')
    on_plane=np.zeros(len(tri),bool)
    for axis in range(3):
        for edge in [-half,half]:on_plane|=np.max(abs(tri[:,:,axis]-edge),axis=1)<=guard_mm
    if np.any(caps&~on_plane):raise ValueError('Clipping-owned face is outside its explicit plane enclosure')
    return owners,caps
