"""Continuous 3D edge-overlap evidence for triangle patch graphs."""
import numpy as np
from rtree.index import Index,Property

def partial_edge_links(edges,labels,tolerance_mm=2e-7):
    edges=np.asarray(edges,float);labels=np.asarray(labels,int)
    if len(edges)!=len(labels):raise ValueError('Edge owner count mismatch')
    if tolerance_mm<0:raise ValueError('Negative numerical identity bound')
    low=edges.min(1)-tolerance_mm;high=edges.max(1)+tolerance_mm
    prop=Property();prop.dimension=3
    tree=Index(((i,tuple(np.r_[low[i],high[i]]),None) for i in range(len(edges))),properties=prop)
    lengths=np.linalg.norm(edges[:,1]-edges[:,0],axis=1);result=[];maximum=0.;checked=0
    for i,(a,b) in enumerate(edges):
        if lengths[i]==0:continue
        ids=np.asarray([j for j in tree.intersection(tuple(np.r_[low[i],high[i]])) if j>i and labels[j]!=labels[i]],dtype=int)
        if not len(ids):continue
        axis=(b-a)/lengths[i];delta=edges[ids]-a;parameters=delta@axis
        perpendicular=delta-parameters[:,:,None]*axis
        error=np.linalg.norm(perpendicular,axis=2).max(1);checked+=len(ids)
        left=np.maximum(0,parameters.min(1));right=np.minimum(lengths[i],parameters.max(1))
        valid=(error<=tolerance_mm)&(right>left)&(lengths[ids]>0)
        for at in np.flatnonzero(valid):
            j=int(ids[at]);maximum=max(maximum,float(error[at]))
            result.append({'edges':[i,j],'components':[int(labels[i]),int(labels[j])],
                'continuous_overlap_length_mm':float(right[at]-left[at]),'maximum_line_separation_mm':float(error[at])})
    tree.close()
    return result,{'candidate_pairs_checked':checked,'verified_partial_edge_links':len(result),'maximum_line_separation_mm':maximum,
        'numerical_identity_bound_mm':tolerance_mm,'criterion':'positive collinear segment overlap in 3D','centre_or_ID_grouping':False}
