"""Conservative plane sections of triangle distance neighbourhoods.

These polygons are mathematical distance certificates, never model geometry.
A triangle's neighbourhood is the union of its normal prism, three finite
edge cylinders and three vertex balls. Circle/ellipse arcs are bounded with
inscribed/circumscribed polygons; no sampling decides absence of a defect.
"""
import numpy as np
import shapely
from shapely.geometry import Polygon,MultiPoint,LineString,Point,GeometryCollection,box

def certificate_shape(shape,tolerance_mm=1e-9):
    if shape.is_empty:return shape
    try:cleaned=shapely.remove_repeated_points(shape,tolerance=tolerance_mm).simplify(tolerance_mm,preserve_topology=True)
    except shapely.errors.GEOSException:return shape
    if cleaned.is_valid and shape.hausdorff_distance(cleaned)<=2*tolerance_mm:return cleaned
    return shape

def certificate_triangulation(polygon,cleanup_mm=1e-9):
    """Validate CDT coverage; normalize only sub-nanometre repeated vertices.

    GEOS can omit a finite triangle next to two almost identical consecutive
    ring points. The certificate guard is much larger than this cleanup bound.
    Model surfaces are never filled using this representation normalization.
    """
    triangles=list(shapely.constrained_delaunay_triangles(polygon).geoms)
    combined=shapely.union_all(triangles);error=polygon.symmetric_difference(combined).area
    budget=max(1e-5,polygon.area*1e-9);normalized=False
    if error>budget:
        clean=shapely.remove_repeated_points(polygon,tolerance=cleanup_mm).simplify(cleanup_mm,preserve_topology=True)
        if not clean.is_valid or clean.geom_type!='Polygon' or len(clean.interiors)!=len(polygon.interiors):raise ValueError('Unsafe distance-certificate ring normalization')
        if polygon.boundary.hausdorff_distance(clean.boundary)>cleanup_mm*2:raise ValueError('Certificate cleanup exceeds geometric bound')
        triangles=list(shapely.constrained_delaunay_triangles(clean).geoms);combined=shapely.union_all(triangles)
        error=polygon.symmetric_difference(combined).area;normalized=True
    if error>budget:raise ValueError('Certified triangulation loses area: '+str(error))
    return triangles,{'coverage_error_mm2':float(error),'normalization_used':normalized,
        'normalization_coordinate_bound_mm':cleanup_mm if normalized else 0.}

def clipped_polygon(points,linear):
    if not len(points):return points
    result=[]
    for i,p in enumerate(points):
        q=points[(i+1)%len(points)];a=float(np.dot(linear[:2],p)+linear[2]);b=float(np.dot(linear[:2],q)+linear[2])
        if a>=0:result.append(p)
        if (a>=0)!=(b>=0):result.append(p+(q-p)*(a/(a-b)))
    return np.asarray(result)

def prism_section(tri,radius):
    normal=np.cross(tri[1]-tri[0],tri[2]-tri[0]);length=np.linalg.norm(normal)
    if not length:return GeometryCollection()
    normal/=length;vertices=np.vstack([tri+radius*normal,tri-radius*normal]);points=[]
    for a,b in [(0,1),(1,2),(2,0),(3,4),(4,5),(5,3),(0,3),(1,4),(2,5)]:
        p,q=vertices[a],vertices[b]
        if p[2]==0:points.append(p[:2])
        if q[2]==0:points.append(q[:2])
        if (p[2]<0)!=(q[2]<0):points.append((p+(q-p)*(p[2]/(p[2]-q[2])))[:2])
    if len(points)<3:return GeometryCollection()
    return MultiPoint(points).convex_hull

def disk(center,radius,outer,segments):
    if radius<=0:return Point(center)
    r=radius/np.cos(np.pi/segments) if outer else radius
    return Point(center).buffer(r,quad_segs=segments//4)

def edge_section(a,b,radius,outer,segments):
    d=b-a;length=np.linalg.norm(d)
    if not length:return GeometryCollection()
    if abs(d[2])<=length*1e-6:
        # This bound remains valid for the very small departure from parallel.
        hmax=max(abs(a[2]),abs(b[2]));hmin=0. if a[2]*b[2]<=0 else min(abs(a[2]),abs(b[2]))
        h=hmin if outer else hmax
        if h>radius:return GeometryCollection()
        r=np.sqrt(max(0.,radius*radius-h*h))
        if outer:r/=np.cos(np.pi/segments)
        return LineString([a[:2],b[:2]]).buffer(r,quad_segs=segments//4)
    # A source plane cuts an infinite edge cylinder in an ellipse. Its
    # quadratic metric has radii R and R*|edge|/|normal component|.
    center=a[:2]-a[2]/d[2]*d[:2];xy=np.linalg.norm(d[:2]);along=d[:2]/xy if xy else np.array([1.,0.]);across=np.array([-along[1],along[0]])
    scale=radius/np.cos(np.pi/segments) if outer else radius
    angles=np.arange(segments)*2*np.pi/segments
    points=center+np.cos(angles)[:,None]*(scale*length/abs(d[2]))*along+np.sin(angles)[:,None]*scale*across
    # Restrict the closest foot to the actual finite segment; endpoint balls
    # are handled separately.
    coefficient=np.r_[d[:2]/(length*length),-float(np.dot(a,d))/(length*length)]
    points=clipped_polygon(points,coefficient)
    points=clipped_polygon(points,np.r_[-coefficient[:2],1-coefficient[2]])
    if len(points)<3:return GeometryCollection()
    # Both the ellipse polygon and the clipping strip are convex. Rebuilding
    # their convex intersection removes duplicate/touching numerical vertices
    # without inventing an unproved support region.
    polygon=MultiPoint(points).convex_hull
    if not polygon.is_valid:raise ValueError('Invalid convex analytic cylinder section')
    return polygon

class TriangleDistanceMasks:
    def __init__(self,vertices,faces):
        self.vertices=np.asarray(vertices,dtype=float);self.faces=np.asarray(faces,dtype=int)

    def masks(self,origin,basis,roi,within_radius,over_radius,segments=128,guard_mm=1e-5):
        if within_radius<0 or over_radius<within_radius:raise ValueError('Invalid distance interval')
        if segments<16 or segments%4:raise ValueError('Arc count must be a multiple of four')
        precision=min(1e-8,guard_mm/100) if guard_mm else 0.
        if precision:
            clean_roi=shapely.set_precision(certificate_shape(roi),precision)
            if not clean_roi.is_empty:roi=clean_roi
        vertices=(self.vertices-origin)@np.asarray(basis).T;tri=vertices[self.faces]
        lo=np.array(roi.bounds[:2]);hi=np.array(roi.bounds[2:]);r=over_radius+guard_mm
        selected=np.flatnonzero((tri[:,:,2].min(1)<=r)&(tri[:,:,2].max(1)>=-r)&
            np.all(tri[:,:,:2].min(1)<=hi+r,axis=1)&np.all(tri[:,:,:2].max(1)>=lo-r,axis=1))
        selected_faces=self.faces[selected]
        crop=box(roi.bounds[0]-2*guard_mm,roi.bounds[1]-2*guard_mm,roi.bounds[2]+2*guard_mm,roi.bounds[3]+2*guard_mm)
        inner=[];outer=[]
        parallel_groups={};nonparallel=[]
        for idx in selected:
            t=tri[idx]
            if np.ptp(t[:,2])<1e-7 and abs(float(np.cross(t[1,:2]-t[0,:2],t[2,:2]-t[0,:2])))>1e-12:
                parallel_groups.setdefault(round(float(t[:,2].mean()),6),[]).append(idx)
            else:nonparallel.append(idx)
        for group in parallel_groups.values():
            tt=tri[group];polys=shapely.polygons(tt[:,:,:2])
            # Nearly coincident source edges can make floating GEOS union
            # cancel a finite overlap while returning a valid polygon. Snap
            # this mathematical support before union, within the guard bound.
            if precision:polys=shapely.set_precision(polys,precision)
            footprint=shapely.union_all(polys)
            if not footprint.is_valid:raise ValueError('Invalid projected target support')
            hmax=float(np.abs(tt[:,:,2]).max());hmin=0. if tt[:,:,2].min()<=0<=tt[:,:,2].max() else float(np.abs(tt[:,:,2]).min())
            for rr,out,collection,h in [(max(0,within_radius-guard_mm),False,inner,hmax),(over_radius+guard_mm,True,outer,hmin)]:
                if h>rr:continue
                distance=np.sqrt(max(0.,rr*rr-h*h))
                if out:distance/=np.cos(np.pi/segments)
                collection.append(footprint.buffer(distance,quad_segs=segments//4).intersection(crop))
        # A union of inner parallel supports is already a continuous proof.
        # Other target primitives can only increase that support; avoid their
        # construction when every point of the measurement domain is covered.
        early_parts=[certificate_shape(p) for p in inner if not p.is_empty]
        if precision:early_parts=shapely.set_precision(early_parts,precision)
        early_inside=certificate_shape(shapely.union_all(early_parts,grid_size=precision))
        if roi.difference(early_inside).is_empty:
            return roi,GeometryCollection(),GeometryCollection(),{
                'method':'complete-inner-parallel-support-v1','continuous_roi_coverage':True,
                'within_radius_mm':within_radius,'over_radius_mm':over_radius,
                'floating_guard_mm':guard_mm,'distance_mask_precision_grid_mm':precision,
                'target_triangles_considered':len(selected),'no_target_geometry_removed':True}
        nonparallel=np.asarray(nonparallel,dtype=int)
        active=self.faces[nonparallel] if len(nonparallel) else np.empty((0,3),int)
        unique_vertices=np.unique(active)
        edges=np.unique(np.sort(np.vstack([active[:,[0,1]],active[:,[1,2]],active[:,[2,0]]]),axis=1),axis=0) if len(active) else []
        for rr,out,collection in [(max(0,within_radius-guard_mm),False,inner),(over_radius+guard_mm,True,outer)]:
            for idx in nonparallel:
                p=prism_section(tri[idx],rr)
                if not p.is_empty:collection.append(p.intersection(crop))
            for idx in unique_vertices:
                p=vertices[idx]
                if abs(p[2])<=rr:
                    collection.append(disk(p[:2],np.sqrt(max(0.,rr*rr-p[2]*p[2])),out,segments).intersection(crop))
            for a,b in edges:
                p=edge_section(vertices[a],vertices[b],rr,out,segments)
                if not p.is_empty:collection.append(p.intersection(crop))
        inner_parts=[certificate_shape(p) for p in inner if not p.is_empty]
        outer_parts=[certificate_shape(p) for p in outer if not p.is_empty]
        if precision:
            inner_parts=shapely.set_precision(inner_parts,precision)
            outer_parts=shapely.set_precision(outer_parts,precision)
        inside=certificate_shape(shapely.union_all(inner_parts,grid_size=precision))
        outside_cover=certificate_shape(shapely.union_all(outer_parts,grid_size=precision))
        if precision:
            inside=shapely.set_precision(inside,precision);outside_cover=shapely.set_precision(outside_cover,precision)
        if not inside.is_valid or not outside_cover.is_valid:raise ValueError('Invalid analytic distance mask union')
        over=roi.difference(outside_cover);within=roi.intersection(inside);uncertain=roi.difference(within).difference(over)
        error=roi.symmetric_difference(shapely.union_all([within,over,uncertain])).area
        if error>max(1e-5,roi.area*1e-9):raise ValueError('Analytic masks do not cover source ROI: area='+str(roi.area)+' error='+str(error))
        return within,over,uncertain,{'method':'triangle-prism-cylinder-ball-plane-section-v2','target_triangles_considered':len(selected),
            'within_radius_mm':within_radius,'over_radius_mm':over_radius,'arc_segments':segments,
            'arc_metric_gap_bound_mm':float((over_radius+guard_mm)*(1/np.cos(np.pi/segments)-1)),
            'floating_guard_mm':guard_mm,'source_roi_area_mm2':roi.area,'coverage_area_error_mm2':error,
            'certificate_representation_roundoff_bound_mm':2e-9,
            'distance_mask_precision_grid_mm':precision,'precision_grid_changes_model_geometry':False,
            'within_area_mm2':within.area,'over_area_mm2':over.area,'uncertain_area_mm2':uncertain.area,
            'new_model_geometry':False}
