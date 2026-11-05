"""Continuous vertical distances between corresponding sloped footprints.

Input footprints are physical surfaces. Clipped polygons are measurement
domains only. Unpaired plan strips never become a height difference.
"""
import numpy as np
import shapely
from shapely.geometry import Polygon,GeometryCollection

def plane_z(triangle):
    t=np.asarray(triangle,float);n=np.cross(t[1]-t[0],t[2]-t[0]);length=np.linalg.norm(n)
    if length==0 or abs(n[2])<=length*1e-12:raise ValueError('No single Z(X,Y)')
    return np.array([-n[0]/n[2],-n[1]/n[2],float(n@t[0])/n[2]])

def values(coeff,points):
    p=np.asarray(points);return p[...,0]*coeff[0]+p[...,1]*coeff[1]+coeff[2]

def clip_positive(shape,coeff,threshold=0.):
    if shape.is_empty:return shape
    a,b,c,d=shape.bounds;ring=np.array([[a-1,b-1],[c+1,b-1],[c+1,d+1],[a-1,d+1]])
    result=[]
    for i,p in enumerate(ring):
        q=ring[(i+1)%4];dp=float(values(coeff,p)-threshold);dq=float(values(coeff,q)-threshold)
        if dp>=0:result.append(p)
        if (dp>=0)!=(dq>=0):result.append(p+(q-p)*dp/(dp-dq))
    return shape.intersection(Polygon(result)) if len(result)>=3 else GeometryCollection()

def nearer(shape,a,b):
    return shapely.union_all([clip_positive(clip_positive(shape,a-b),-(a+b)),
        clip_positive(clip_positive(shape,-(a-b)),a+b)])

def measure(source_shape,source_plane,candidates,tolerance,source_error=0.):
    if tolerance<0 or source_error<0:raise ValueError('Negative distance bound')
    active=[]
    for c in candidates:
        slope=np.asarray(c['plane'])[:2]
        if float(source_plane[:2]@slope)<=0:continue
        shape=source_shape.intersection(c['shape'])
        if shape.area>0:active.append({**c,'shape':shape,'difference':np.asarray(c['plane'])-source_plane})
    paired=shapely.union_all([c['shape'] for c in active]) if active else GeometryCollection()
    rows=[];assigned=GeometryCollection()
    for i,c in enumerate(active):
        region=c['shape']
        for j,other in enumerate(active):
            if i==j:continue
            overlap=region.intersection(other['shape'])
            if overlap.is_empty:continue
            region=shapely.union_all([region.difference(overlap),nearer(overlap,c['difference'],other['difference'])])
        # Equal-distance ties have one owner, without duplicating area.
        region=region.difference(assigned)
        if region.is_empty or region.area==0:continue
        assigned=shapely.union_all([assigned,region])
        slope_norm=float(np.linalg.norm(np.asarray(c['plane'])[:2]))
        err=source_error*(1+slope_norm)+float(c.get('surface_error',0))*np.sqrt(1+slope_norm*slope_norm)
        d=c['difference'];limit=tolerance+err
        over=shapely.union_all([clip_positive(region,d,limit+1e-8),clip_positive(region,-d,limit+1e-8)])
        lower=max(0.,tolerance-err)
        within=clip_positive(clip_positive(region,d,-lower),-d,-lower)
        uncertain=region.difference(over).difference(within)
        points=shapely.get_coordinates(region)
        rows.append({'target':c.get('key'),'shape':region,'over':over,'within':within,'uncertain':uncertain,
            'difference':d,'uncertainty_mm':err,'maximum_vertical_mm':float(np.abs(values(d,points)).max()),
            'minimum_vertical_mm':float(np.abs(values(d,points)).min())})
    lost=paired.symmetric_difference(assigned).area
    if lost>max(1e-5,source_shape.area*1e-9):raise ValueError('Vertical nearest-surface partition loses coverage')
    return rows,source_shape.difference(paired),{'source_area_mm2':float(source_shape.area),'paired_area_mm2':float(paired.area),
        'unpaired_plan_area_mm2':float(source_shape.difference(paired).area),'partition_error_mm2':float(lost)}
