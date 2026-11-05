"""Independent solid-angle winding for oriented, closed input triangles.

An outside decision additionally requires a verified positive distance from
the source boundary. A winding value alone never certifies boundary points.
"""
import numpy as np

def winding_numbers(points,vertices,faces,point_chunk=32,triangle_chunk=2048):
    points=np.asarray(points,float);tri=np.asarray(vertices,float)[np.asarray(faces,int)]
    if not np.isfinite(points).all() or not np.isfinite(tri).all():raise ValueError('Nonfinite winding input')
    if point_chunk<1 or triangle_chunk<1:raise ValueError('Positive chunk sizes required')
    result=np.zeros(len(points))
    for start in range(0,len(points),point_chunk):
        p=points[start:start+point_chunk];total=np.zeros(len(p));compensation=np.zeros(len(p))
        for offset in range(0,len(tri),triangle_chunk):
            t=tri[offset:offset+triangle_chunk]
            a=t[None,:,0]-p[:,None];b=t[None,:,1]-p[:,None];c=t[None,:,2]-p[:,None]
            la=np.linalg.norm(a,axis=2);lb=np.linalg.norm(b,axis=2);lc=np.linalg.norm(c,axis=2)
            if np.any(la==0)|np.any(lb==0)|np.any(lc==0):raise ValueError('Winding at a source vertex is undefined')
            numerator=np.einsum('ijk,ijk->ij',a,np.cross(b,c))
            denominator=la*lb*lc+np.einsum('ijk,ijk->ij',a,b)*lc+np.einsum('ijk,ijk->ij',b,c)*la+np.einsum('ijk,ijk->ij',c,a)*lb
            angles=2*np.arctan2(numerator,denominator)
            value=angles.sum(axis=1)-compensation;new=total+value
            compensation=(new-total)-value;total=new
        result[start:start+len(p)]=total/(4*np.pi)
    return result
