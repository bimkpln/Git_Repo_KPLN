"""A common rigid frame for operands and computational clipping boxes."""
import numpy as np
def rotation_matrix(degrees):
    x,y,z=np.radians(np.asarray(degrees,float));cx,cy,cz=np.cos([x,y,z]);sx,sy,sz=np.sin([x,y,z])
    rx=np.array([[1,0,0],[0,cx,-sx],[0,sx,cx]]);ry=np.array([[cy,0,sy],[0,1,0],[-sy,0,cy]]);rz=np.array([[cz,-sz,0],[sz,cz,0],[0,0,1]])
    R=rz@ry@rx
    if np.max(abs(R@R.T-np.eye(3)))>1e-12 or abs(np.linalg.det(R)-1)>1e-12:raise ValueError('Non-rigid numerical frame')
    return R
def independent_volume(native):
    mesh=native.to_mesh64();v=np.asarray(mesh.vert_properties)[:,:3];f=np.asarray(mesh.tri_verts,dtype=np.int64)
    if not len(f):return 0.
    center=(v.min(0)+v.max(0))/2;t=v[f]-center
    return float(np.einsum('ij,ij->i',t[:,0],np.cross(t[:,1],t[:,2])).sum()/6)
