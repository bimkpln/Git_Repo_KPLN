"""Native directed contours, preserving small holes and original segments."""
import numpy as np
import shapely
from fractions import Fraction
from shapely.geometry import Polygon,MultiPolygon,GeometryCollection
from shapely.affinity import translate
from section_geometry import section_from_loops as legacy_section

def _polygons(shape):
    if shape.geom_type=='Polygon':return [shape] if not shape.is_empty else []
    return [p for g in getattr(shape,'geoms',[]) for p in _polygons(g)]

def _signed_area(points):
    p=np.asarray(points,float);q=p-(p.min(0)+p.max(0))/2;nextp=np.roll(q,-1,axis=0)
    products=q[:,0]*nextp[:,1]-q[:,1]*nextp[:,0];area=float(products.sum()/2)
    if abs(area)<=np.finfo(float).eps*sum(abs(products))*16:
        a=[[Fraction(float(v)) for v in row] for row in p]
        area=float(sum(a[i][0]*a[(i+1)%len(a)][1]-a[(i+1)%len(a)][0]*a[i][1] for i in range(len(a)))/2)
    return area

def native_section_from_loops(loops,expected_area,allow_noded_contours=False,representation_bound_mm=1e-7):
    if expected_area is None:raise ValueError('Native directed contour area is required')
    positive=[];negative=[];noded=[];zero=[];signed_total=0.
    for index,loop in enumerate(loops):
        points=np.asarray(loop,float)
        if len(points)<3:zero.append(index);continue
        signed=_signed_area(points);signed_total+=signed;raw=Polygon(points)
        if signed==0:
            if raw.area>0:raise ValueError('Mixed winding native contour has zero signed area')
            zero.append(index);continue
        if raw.is_valid:parts=[raw]
        else:
            if not allow_noded_contours:raise ValueError('Native contour requires original-segment arrangement')
            origin=(points.min(0)+points.max(0))/2
            part,proof=legacy_section([points-origin],abs(signed),allow_noded_contours=True)
            parts=_polygons(translate(part,*origin));noded.append({'contour':index,'proof':proof})
        if not parts:raise ValueError('Positive-area native contour disappeared')
        if any(not p.is_valid for p in parts):raise ValueError('Local contour arrangement is invalid')
        (positive if signed>0 else negative).extend(parts)
    budget=max(1e-4,abs(expected_area)*1e-9)
    if abs(signed_total-expected_area)>budget:raise ValueError('Directed native contour areas disagree with native material')
    # Negative rings may enclose positive material islands. Keep their actual
    # contours explicitly; no global line-network cancellation is permitted.
    for p in negative:
        positive.extend(Polygon(r) for r in p.interiors)
    outers=[Polygon(p.exterior) for p in positive];tree=shapely.STRtree(outers)
    holes={i:list(p.interiors) for i,p in enumerate(positive)};parents=[]
    for p in negative:
        candidates=tree.query(Polygon(p.exterior),predicate='covered_by')
        if not len(candidates):raise ValueError('Native negative contour has no enclosing material contour')
        owner=int(min(candidates,key=lambda i:outers[int(i)].area));holes[owner].append(p.exterior);parents.append(owner)
    built=[Polygon(p.exterior,holes[i]) for i,p in enumerate(positive)]
    if any(not p.is_valid for p in built):raise ValueError('Native shell/hole nesting is invalid')
    shape=built[0] if len(built)==1 else MultiPolygon(built) if built else GeometryCollection()
    if not shape.is_valid:shape=shapely.union_all(built)
    if not shape.is_valid:raise ValueError('Native contour material union is invalid')
    area_error=abs(shape.area-expected_area)
    if area_error>budget:raise ValueError('Native contour reconstruction loses material area: '+str(area_error))
    enclosure=shape.buffer(representation_bound_mm) if representation_bound_mm else shape
    omitted=0.
    for p in built:
        missing=p.difference(shape)
        if missing.area==p.area and p.area>0:raise ValueError('A complete positive native contour was omitted')
        if missing.difference(enclosure).area>0:raise ValueError('Native material omission exceeds representation distance')
        omitted+=missing.area
    # Every native negative contour is checked, even when its area is far
    # below the global numerical area-balance budget.
    hole_proofs=[]
    for p,owner in zip(negative,parents):
        islands=[built[int(i)] for i in tree.query(Polygon(p.exterior),predicate='contains') if int(i)!=owner]
        void=p.difference(shapely.union_all(islands)) if islands else p
        filled=void.intersection(shape).area
        if filled>0:raise ValueError('Native negative contour was filled during reconstruction')
        hole_proofs.append({'native_void_area_mm2':p.area,'unexpected_material_area_mm2':filled})
    return shape,{'method':'native-directed-per-contour-original-segment-arrangement-v1','contours':len(loops),
        'expected_native_area_mm2':float(expected_area),'reconstructed_area_mm2':shape.area,'area_error_mm2':area_error,
        'noded_contours':noded,'zero_signed_area_contours':zero,'native_holes_preserved':hole_proofs,
        'omitted_material_boundary_area_mm2':omitted,'representation_bound_mm':representation_bound_mm,
        'no_hole_size_filter':True,'new_or_guessed_contours':0,'model_geometry_modified':False}
