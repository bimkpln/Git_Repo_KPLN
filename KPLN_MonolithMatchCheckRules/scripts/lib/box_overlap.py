"""Exact SAT for positive-volume intersection of two full-dimensional boxes.

Each domain is defined by low < basis @ point < high. The basis may be
oblique; its exact binary input values are used, without an orthogonalization
or a geometric size threshold. These are mathematical domains, not model
solids or proof that the corresponding holes have the same purpose.
"""
from fractions import Fraction as F
from itertools import product
from numbers import Integral

def rational(x):
    return x if isinstance(x,F) else F(int(x)) if isinstance(x,Integral) else F(float(x))

def dot(a,b):return sum(x*y for x,y in zip(a,b))
def cross(a,b):return tuple(a[(i+1)%3]*b[(i+2)%3]-a[(i+2)%3]*b[(i+1)%3] for i in range(3))

def domain(basis,bounds):
    if len(basis)!=3 or any(len(r)!=3 for r in basis) or len(bounds)!=2 or any(len(r)!=3 for r in bounds):
        raise ValueError('Expected a 3x3 basis and two 3D bounds')
    M=[tuple(map(rational,r)) for r in basis];low,high=[tuple(map(rational,r)) for r in bounds]
    if any(a>=b for a,b in zip(low,high)):raise ValueError('The domain must have positive extent in every direction')
    columns=[cross(M[1],M[2]),cross(M[2],M[0]),cross(M[0],M[1])];det=dot(M[0],columns[0])
    if not det:raise ValueError('Singular domain basis')
    edges=[tuple(x/det for x in col) for col in columns]
    vertices=[tuple(sum(edges[i][j]*q[i] for i in range(3)) for j in range(3)) for q in product(*zip(low,high))]
    return {'basis':M,'edges':edges,'vertices':vertices}

def exact_box_overlap(first_basis,first_bounds,second_basis,second_bounds):
    a=domain(first_basis,first_bounds);b=domain(second_basis,second_bounds)
    axes=a['basis']+b['basis']+[cross(u,v) for u in a['edges'] for v in b['edges']]
    count=0
    for axis in axes:
        if not any(axis):continue
        count+=1;pa=[dot(axis,v) for v in a['vertices']];pb=[dot(axis,v) for v in b['vertices']]
        overlap=min(max(pa),max(pb))-max(min(pa),min(pb))
        if overlap<=0:
            return {'positive_volume_overlap':False,'method':'exact-rational-parallelepiped-separating-axes-v1',
                    'separating_axis':[str(x) for x in axis],'projected_overlap':str(overlap),
                    'touching_on_this_axis':overlap==0,'tested_nonzero_axes':count,'physical_size_cutoff':False}
    return {'positive_volume_overlap':True,'method':'exact-rational-parallelepiped-separating-axes-v1',
            'tested_nonzero_axes':count,'all_required_face_and_edge_cross_axes_tested':True,'physical_size_cutoff':False}
