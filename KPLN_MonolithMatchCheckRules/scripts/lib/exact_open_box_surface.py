"""Open convex query-box/source-surface predicates, with rational fallback.

Boundary-only contact is allowed; any original facet meeting the box interior
is retained. No material conclusion follows without a closed-source membership
test. Floats only cull faces with a conservative arithmetic separation bound.
"""
from fractions import Fraction as F
import numpy as np

def clip_polygon(poly,axis,bound,lower):
    if not poly:return []
    out=[]
    for a,b in zip(poly,poly[1:]+poly[:1]):
        da=(a[axis]-bound) if lower else (bound-a[axis]);db=(b[axis]-bound) if lower else (bound-b[axis])
        if da>=0:out.append(a)
        if (da>=0)!=(db>=0):
            t=da/(da-db);out.append(tuple(a[k]+t*(b[k]-a[k]) for k in range(3)))
    return out

def triangle_intersects_open_box(triangle,low,high):
    poly=[tuple(F(x) for x in p) for p in triangle];low=tuple(map(F,low));high=tuple(map(F,high))
    if any(a>=b for a,b in zip(low,high)):raise ValueError('A positive query box is required')
    for axis in range(3):
        poly=clip_polygon(poly,axis,low[axis],True);poly=clip_polygon(poly,axis,high[axis],False)
        if not poly:return False
    # The mean is strictly interior iff the clipped polygon is not wholly on
    # any box support plane. This also handles edge/point source contacts.
    return all(any(p[k]>low[k] for p in poly) and any(p[k]<high[k] for p in poly) for k in range(3))

def source_surface_intersects_open_box(vertices,faces,matrix,low,high):
    vertices=np.asarray(vertices);faces=np.asarray(faces);matrix=np.asarray(matrix)
    low=np.asarray(low);high=np.asarray(high)
    if np.any(low>=high):raise ValueError('A positive query box is required')
    framed=vertices@matrix.T;tri=framed[faces]
    errors=(abs(vertices)@abs(matrix).T)*np.finfo(float).eps*32+1e-8;err=errors[faces].max(1)
    separated=np.any((tri.max(1)<=low-err)|(tri.min(1)>=high+err),axis=1)
    candidates=np.flatnonzero(~separated);rational={};M=[[F(float(x)) for x in row] for row in matrix]
    def point(i):
        i=int(i)
        if i not in rational:
            p=[F(float(x)) for x in vertices[i]];rational[i]=tuple(sum(a*b for a,b in zip(row,p)) for row in M)
        return rational[i]
    for index in candidates:
        if triangle_intersects_open_box([point(i) for i in faces[index]],list(map(float,low)),list(map(float,high))):
            return {'intersects':True,'source_face_index':int(index),'faces_considered':len(faces),'exact_fallback_candidates':len(candidates)}
    return {'intersects':False,'faces_considered':len(faces),'exact_fallback_candidates':len(candidates),
            'all_source_faces_covered':True,'boundary_only_contact_not_called_interior':True}
