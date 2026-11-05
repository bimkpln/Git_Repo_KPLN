"""Independent parent coverage by bounded child measurement domains."""
import numpy as np
import shapely
from shapely.geometry import MultiPoint
def parent_coverage(parent,children,guard_mm):
    if guard_mm<=0:raise ValueError('Positive numerical coverage bound required')
    t=np.asarray(parent,float);child=np.asarray(children,float).reshape(-1,3,3)
    if not len(child):return {'passed':False,'reason':'no_child_certificates'}
    edges=t[[1,2,0]]-t;longest=int(np.argmax(np.linalg.norm(edges,axis=1)));u=edges[longest];length=np.linalg.norm(u)
    if length==0:u=np.array([1.,0.,0.])
    else:u=u/length
    n=np.cross(t[1]-t[0],t[2]-t[0]);size=np.linalg.norm(n)
    if size==0:
        basis=np.eye(3)[int(np.argmin(abs(u)))];n=np.cross(u,basis);n/=np.linalg.norm(n)
    else:n/=size
    v=np.cross(n,u);vl=np.linalg.norm(v)
    if vl<1e-8:
        axis=np.eye(3)[int(np.argmin(abs(u)))];v=np.cross(axis,u);vl=np.linalg.norm(v)
    v/=vl;n=np.cross(u,v);n/=np.linalg.norm(n);v=np.cross(n,u);v/=np.linalg.norm(v)
    B=np.array([u,v,n]);source=(t-t[0])@B.T;actual=(child-t[0])@B.T
    normal_error=float(abs(actual[:,:,2]).max());parent_normal_error=float(abs(source[:,2]).max())
    source_shape=MultiPoint(source[:,:2]).convex_hull
    projected=shapely.convex_hull(shapely.polygons(actual[:,:,:2]));cover=shapely.union_all(shapely.buffer(projected,guard_mm/2,quad_segs=4))
    passed=normal_error+parent_normal_error<=guard_mm/2 and cover.covers(source_shape)
    return {'passed':bool(passed),'maximum_child_normal_error_mm':normal_error,'parent_normal_error_mm':parent_normal_error,
        'whole_parent_covered_by_child_neighbourhoods':bool(cover.covers(source_shape)),
        'distance_transfer_bound_mm':guard_mm,'uses_1_Lipschitz_distance':True,'child_domains_are_not_model_geometry':True}
