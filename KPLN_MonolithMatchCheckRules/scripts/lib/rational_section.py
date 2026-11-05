"""Exact rational intersections of an indexed closed surface and a plane.

No coordinate welding, guessed caps, polygon repair, or size filtering.
Directed cycles follow the original indexed triangle topology.
"""
from fractions import Fraction as F
import numpy as np

def section_cycles(vertices,faces,matrix,offset,positive_side_limit=False):
    M=[[F(float(x)) for x in row] for row in matrix];v=[[F(float(x)) for x in row] for row in vertices];offset=F(float(offset))
    p=[[sum(M[a][b]*q[b] for b in range(3)) for a in range(3)] for q in v]
    d=[q[2]-offset for q in p]
    on_vertex=any(x==0 for x in d)
    if on_vertex and not positive_side_limit:raise ValueError('Cut plane contains an input vertex; one-sided treatment required')
    signs=[1 if x>0 else -1 for x in d]
    determinant=sum(M[0][i]*(M[1][(i+1)%3]*M[2][(i+2)%3]-M[1][(i+2)%3]*M[2][(i+1)%3]) for i in range(3))
    if not determinant:raise ValueError('Singular section coordinates')
    points={};outgoing={};incoming={}
    def crossing(i,j):
        key=tuple(sorted((int(i),int(j))))
        if key not in points:
            a,b=key;t=d[a]/(d[a]-d[b]);points[key]=tuple(p[a][k]+t*(p[b][k]-p[a][k]) for k in range(2))
        return key
    for face in faces:
        start=end=None
        for i,j in zip(face,[face[1],face[2],face[0]]):
            if signs[int(i)]>0>signs[int(j)]:start=crossing(i,j)
            if signs[int(i)]<0<signs[int(j)]:end=crossing(i,j)
        if start is None and end is None:continue
        if start is None or end is None:raise ValueError('Incomplete exact triangle intersection')
        if determinant<0:start,end=end,start
        if start in outgoing or end in incoming:raise ValueError('Source section topology is not a directed manifold')
        outgoing[start]=end;incoming[end]=start
    if set(outgoing)!=set(incoming):raise ValueError('Source section has an open exact contour')
    unseen=set(outgoing);loops=[]
    while unseen:
        first=min(unseen);key=first;loop=[]
        while True:
            if key not in unseen:raise ValueError('Exact contour joins another cycle')
            unseen.remove(key);q=points[key]
            if not loop or q!=loop[-1]:loop.append(q)
            key=outgoing[key]
            if key==first:break
        if len(loop)>1 and loop[0]==loop[-1]:loop.pop()
        # Exact area is a topological sign test. Zero is recorded explicitly.
        area=sum(loop[i][0]*loop[(i+1)%len(loop)][1]-loop[(i+1)%len(loop)][0]*loop[i][1] for i in range(len(loop)))/2 if loop else F(0)
        loops.append({'points':loop,'signed_area':area})
    return loops,{'method':'exact-rational-indexed-source-plane-intersection-v1','intersected_input_edges':len(points),
        'cycles':len(loops),'positive_side_limit_applied':bool(on_vertex and positive_side_limit),'coordinate_welding':False,'new_caps':0,'size_filter':False}
