"""Conservative Euclidean enclosure of a parent by child triangles."""
import numpy as np
import shapely
from shapely.geometry import MultiPoint

def parent_coverage(parent,children,guard_mm):
    if guard_mm<=0:raise ValueError('Positive coverage distance required')
    t=np.asarray(parent,float);children=np.asarray(children,float).reshape(-1,3,3)
    if not len(children):return {'passed':False,'reason':'no_child_domains'}
    edges=t[[1,2,0]]-t;u=edges[np.argmax(np.linalg.norm(edges,axis=1))];length=np.linalg.norm(u)
    u=u/length if length else np.array([1.,0.,0.])
    normal=np.cross(t[1]-t[0],t[2]-t[0]);v=np.cross(normal,u);size=np.linalg.norm(v)
    if size<1e-20:v=np.cross(np.eye(3)[np.argmin(abs(u))],u);size=np.linalg.norm(v)
    v/=size;n=np.cross(u,v);n/=np.linalg.norm(n);v=np.cross(n,u)
    B=np.array([u,v,n]);p=(t-t[0])@B.T;q=(children-t[0])@B.T
    parent_height=float(abs(p[:,2]).max());heights=abs(q[:,:,2]).max(1)+parent_height
    # For each projected child point choose the corresponding barycentric
    # point in 3D. Its height is bounded by its triangle's vertex heights.
    # Inscribed polygonal buffers are conservative Euclidean disks.
    radii=np.sqrt(np.maximum(0.,guard_mm**2-heights**2))
    pieces=[]
    for triangle,h,r in zip(q,heights,radii):
        if h>=guard_mm:continue
        shape=MultiPoint(triangle[:,:2]).convex_hull
        pieces.append(shape.buffer(max(0.,r-1e-10),quad_segs=16))
    cover=shapely.union_all(pieces);source=MultiPoint(p[:,:2]).convex_hull
    return {'passed':bool(cover.covers(source)),'whole_parent_covered':bool(cover.covers(source)),
            'distance_enclosure_mm':guard_mm,'parent_normal_error_mm':parent_height,
            'child_normal_errors_mm':heights.tolist(),'child_projected_radii_mm':radii.tolist(),
            'proof':'per-child Pythagorean distance enclosure with inscribed disk buffers',
            'new_model_geometry':False}
