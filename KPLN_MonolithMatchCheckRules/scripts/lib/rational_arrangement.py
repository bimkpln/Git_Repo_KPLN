"""Exact noding and winding of directed planar source segments."""
from fractions import Fraction as F
from collections import defaultdict
import numpy as np
import shapely
from shapely.geometry import LineString,box,GeometryCollection

def cross(a,b):return a[0]*b[1]-a[1]*b[0]
def sub(a,b):return (a[0]-b[0],a[1]-b[1])
def at(a,r,t):return (a[0]+r[0]*t,a[1]+r[1]*t)

def positive_region(cycles):
    original=[]
    for loop in cycles:
        for a,b in zip(loop,loop[1:]+loop[:1]):
            a=tuple(a);b=tuple(b)
            if a!=b:original.append((a,b))
    bounds=[]
    for a,b in original:
        lo=np.minimum(list(map(float,a)),list(map(float,b)));hi=np.maximum(list(map(float,a)),list(map(float,b)))
        lo=np.nextafter(lo,-np.inf);hi=np.nextafter(hi,np.inf);bounds.append(box(*lo,*hi))
    tree=shapely.STRtree(bounds);cuts=[{F(0),F(1)} for _ in original];crossings=0
    for i,(a,b) in enumerate(original):
        r=sub(b,a)
        for j in tree.query(bounds[i]):
            j=int(j)
            if j<=i:continue
            c,d=original[j];s=sub(d,c);den=cross(r,s);ca=sub(c,a)
            if den:
                t=cross(ca,s)/den;u=cross(ca,r)/den
                if 0<=t<=1 and 0<=u<=1:cuts[i].add(t);cuts[j].add(u);crossings+=1
            elif cross(ca,r)==0:
                axis=0 if r[0] else 1;other_axis=0 if s[0] else 1
                for p in [c,d]:
                    t=(p[axis]-a[axis])/r[axis]
                    if 0<=t<=1:cuts[i].add(t)
                for p in [a,b]:
                    u=(p[other_axis]-c[other_axis])/s[other_axis]
                    if 0<=u<=1:cuts[j].add(u)
    weights=defaultdict(int)
    for (a,b),params in zip(original,cuts):
        r=sub(b,a);values=sorted(params)
        for t,u in zip(values,values[1:]):
            p=at(a,r,t);q=at(a,r,u)
            key=(p,q) if p<q else(q,p);weights[key]+=1 if p<q else -1
    edges=[e for e,w in weights.items() if w];vertices={p for e in edges for p in e};float_to_exact={};maximum_error=0.
    for p in vertices:
        value=tuple(map(float,p))
        if value in float_to_exact and float_to_exact[value]!=p:raise ValueError('Distinct exact nodes cannot be represented by separate floating points')
        float_to_exact[value]=p;maximum_error=max(maximum_error,max(float(abs(F(v)-x)) for v,x in zip(value,p)))
    graph=defaultdict(list)
    for a,b in edges:graph[a].append(b);graph[b].append(a)
    unseen=set(vertices);components=0
    while unseen:
        stack=[unseen.pop()];components+=1
        while stack:
            for q in graph[stack.pop()]:
                if q in unseen:unseen.remove(q);stack.append(q)
    lines=[LineString([list(map(float,a)),list(map(float,b))]) for a,b in edges]
    cells=list(shapely.polygonize(lines).geoms)
    expected_faces=len(edges)-len(vertices)+components
    if len(cells)!=expected_faces:raise ValueError('Floating polygonization does not preserve the exact planar graph')
    selected=[];free=[];windings=[]
    for cell in cells:
        if not cell.is_valid:raise ValueError('Exact-noded planar face is not representable as a valid polygon')
        ring=[float_to_exact[tuple(p)] for p in list(cell.exterior.coords)[:-1]]
        signed=sum(cross(ring[i],ring[(i+1)%len(ring)]) for i in range(len(ring)))
        if signed==0:raise ValueError('Polygonization created a zero exact-area face')
        a,b=ring[0],ring[1];direction=sub(b,a);normal=(-direction[1],direction[0])
        if signed<0:normal=(-normal[0],-normal[1])
        midpoint=((a[0]+b[0])/2,(a[1]+b[1])/2);hits=[]
        for c,d in edges:
            s=sub(d,c);den=cross(normal,s)
            if not den:continue
            delta=sub(c,midpoint);t=cross(delta,s)/den;u=cross(delta,normal)/den
            if t>0 and 0<=u<=1:hits.append(t)
        if not hits:raise ValueError('Bounded planar face has no exact interior witness')
        point=at(midpoint,normal,min(hits)/2);winding=0
        for a,b in original:
            orient=cross(sub(b,a),sub(point,a))
            if a[1]<=point[1]<b[1] and orient>0:winding+=1
            if b[1]<=point[1]<a[1] and orient<0:winding-=1
        windings.append(winding)
        (selected if winding>0 else free).append(cell)
    result=shapely.coverage_union_all(selected) if selected else GeometryCollection()
    if not result.is_valid:raise ValueError('Exact planar material coverage is invalid after floating representation')
    if any(p.intersection(result).area>0 for p in free):raise ValueError('An exact free face was filled')
    if any(p.difference(result).area>0 for p in selected):raise ValueError('An exact material face was omitted')
    return result,{'method':'exact-rational-noding-positive-winding-and-verified-face-coverage-v1',
        'source_segments':len(original),'atomic_edges':len(edges),'exact_nodes':len(vertices),'graph_components':components,
        'expected_bounded_faces':expected_faces,'represented_faces':len(cells),'face_windings':windings,
        'material_faces':len(selected),'free_faces':len(free),'maximum_coordinate_representation_error_mm':maximum_error,
        'size_filter':False,'new_or_guessed_edges':0}
