"""Orthonormal measurement frames and affine lifting onto original triangles."""
import numpy as np

def measurement_frame(triangles):
    t=np.asarray(triangles,float)
    edges=t[:,[1,2,0]]-t
    lengths=np.linalg.norm(edges,axis=2)
    cross=np.cross(t[:,1]-t[:,0],t[:,2]-t[:,0])
    quality=np.linalg.norm(cross,axis=1)/np.maximum(lengths.max(1)**2,np.finfo(float).tiny)
    i=int(np.argmax(quality))
    if quality[i]<=128*np.finfo(float).eps:
        raise ValueError('Source triangle has numerically indeterminate plane')
    u=edges[i,int(np.argmax(lengths[i]))];u=u/np.linalg.norm(u)
    n=cross[i]/np.linalg.norm(cross[i]);n-=u*(n@u);n/=np.linalg.norm(n)
    v=np.cross(n,u);v/=np.linalg.norm(v)
    n=np.cross(u,v)
    return t[i,0].copy(),np.asarray([u,v,n])

def lift_triangle(xy,original,origin,basis):
    projected=(np.asarray(original)-origin)@basis[:2].T
    matrix=(projected[1:]-projected[0]).T
    if np.linalg.cond(matrix)>1e12:
        raise ValueError('Source triangle projection is ill conditioned')
    weights=np.linalg.solve(matrix,(np.asarray(xy)-projected[0]).T).T
    bary=np.c_[1-weights.sum(1),weights]
    if np.min(bary)<-1e-8 or np.max(bary)>1+1e-8:
        raise ValueError('Partition leaves original source triangle')
    # A proof uses the original points; no plane-normal reconstruction.
    points=bary@np.asarray(original)
    error=float(np.max(np.linalg.norm((points-origin)@basis[:2].T-xy,axis=1)))
    if error>1e-6:raise ValueError('Affine source lift exceeds numerical guard')
    return points,error
