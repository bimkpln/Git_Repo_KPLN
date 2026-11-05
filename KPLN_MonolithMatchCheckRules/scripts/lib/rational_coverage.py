"""Exact planar coverage predicates without floating polygonization or welding."""
from fractions import Fraction as F
from collections import defaultdict
from functools import cmp_to_key
import numpy as np
import shapely
from shapely.geometry import box
from rational_arrangement import cross,sub,at

def loops(shape):
    result=[]
    parts=[shape] if shape.geom_type=='Polygon' else list(getattr(shape,'geoms',[]))
    for p in parts:
        if p.geom_type!='Polygon':continue
        p=shapely.orient_polygons(p,exterior_cw=False)
        result.append([tuple(F(float(x)) for x in q) for q in list(p.exterior.coords)[:-1]])
        result.extend([tuple(F(float(x)) for x in q) for q in list(r.coords)[:-1]] for r in p.interiors)
    return result

def coverage(query,materials,material_cycles=None):
    original=[]
    for label,cycles in [(0,loops(query)),(1,material_cycles if material_cycles is not None else [r for p in materials for r in loops(p)])]:
        for loop in cycles:
            original.extend((a,b,label) for a,b in zip(loop,loop[1:]+loop[:1]) if a!=b)
    bounds=[]
    for a,b,_ in original:
        low=np.nextafter(np.minimum(list(map(float,a)),list(map(float,b))),-np.inf)
        high=np.nextafter(np.maximum(list(map(float,a)),list(map(float,b))),np.inf)
        bounds.append(box(*low,*high))
    tree=shapely.STRtree(bounds);cuts=[{F(0),F(1)} for _ in original]
    for i,(a,b,_) in enumerate(original):
        r=sub(b,a)
        for j in tree.query(bounds[i]):
            j=int(j)
            if j<=i:continue
            c,d,_=original[j];s=sub(d,c);den=cross(r,s);ca=sub(c,a)
            if den:
                t=cross(ca,s)/den;u=cross(ca,r)/den
                if 0<=t<=1 and 0<=u<=1:cuts[i].add(t);cuts[j].add(u)
            elif cross(ca,r)==0:
                axis=0 if r[0] else 1;other=0 if s[0] else 1
                for p in [c,d]:
                    t=(p[axis]-a[axis])/r[axis]
                    if 0<=t<=1:cuts[i].add(t)
                for p in [a,b]:
                    u=(p[other]-c[other])/s[other]
                    if 0<=u<=1:cuts[j].add(u)
    weights=defaultdict(lambda:[0,0])
    for (a,b,label),params in zip(original,cuts):
        values=sorted(params);r=sub(b,a)
        for t,u in zip(values,values[1:]):
            p=at(a,r,t);q=at(a,r,u);edge=(p,q) if p<q else(q,p)
            weights[edge][label]+=1 if p<q else -1
    edges=[e for e,w in weights.items() if any(w)];graph=defaultdict(list)
    for a,b in edges:graph[a].append(b);graph[b].append(a)
    def compare(origin,a,b):
        u=sub(a,origin);v=sub(b,origin)
        hu=0 if u[1]>0 or (u[1]==0 and u[0]>=0) else 1
        hv=0 if v[1]>0 or (v[1]==0 and v[0]>=0) else 1
        if hu!=hv:return -1 if hu<hv else 1
        c=cross(u,v)
        if c:return -1 if c>0 else 1
        raise ValueError('Exact noding left overlapping rays')
    positions={}
    for p,neighbours in graph.items():
        neighbours.sort(key=cmp_to_key(lambda a,b:compare(p,a,b)))
        positions[p]={q:i for i,q in enumerate(neighbours)}
    unseen={(a,b) for a,b in edges}|{(b,a) for a,b in edges};cells=0;query_cells=0;uncovered=[];material_cells=0;material_windings=[]
    while unseen:
        first=min(unseen);edge=first;ring=[]
        while True:
            if edge not in unseen:raise ValueError('Exact half-edge traversal joins a different cycle')
            unseen.remove(edge);a,b=edge;ring.append(a)
            neighbours=graph[b];edge=(b,neighbours[(positions[b][a]-1)%len(neighbours)])
            if edge==first:break
        signed=sum(cross(ring[i],ring[(i+1)%len(ring)]) for i in range(len(ring)))
        if signed==0:continue
        cells+=1;a,b=ring[0],ring[1];r=sub(b,a);normal=(-r[1],r[0]);mid=((a[0]+b[0])/2,(a[1]+b[1])/2);hits=[]
        for c,d in edges:
            s=sub(d,c);den=cross(normal,s)
            if not den:continue
            delta=sub(c,mid);t=cross(delta,s)/den;u=cross(delta,normal)/den
            if t>0 and 0<=u<=1:hits.append(t)
        point=at(mid,normal,min(hits)/2 if hits else F(1));w=[0,0]
        for a,b,label in original:
            orient=cross(sub(b,a),sub(point,a))
            if a[1]<=point[1]<b[1] and orient>0:w[label]+=1
            if b[1]<=point[1]<a[1] and orient<0:w[label]-=1
        if any(v<0 for v in w):raise ValueError('Valid polygon produced negative winding')
        if w[0]>0:
            query_cells+=1;material_windings.append(w[1])
            if w[1]>0:material_cells+=1
            else:uncovered.append([str(x) for x in point])
    if query.area>0 and not query_cells:raise ValueError('Positive query lost all exact faces')
    return {'method':'exact-rational-half-edge-coverage-v1','covered':not uncovered,'uncovered_exact_witnesses':uncovered,
        'atomic_edges':len(edges),'half_edge_cycles':cells,'query_face_cycles':query_cells,'covered_query_face_cycles':material_cells,'minimum_material_winding_in_query':min(material_windings,default=0),'maximum_material_winding_in_query':max(material_windings,default=0),
        'floating_polygonization':False,'coordinate_welding':False,'size_filter':False}
