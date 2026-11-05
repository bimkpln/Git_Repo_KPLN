"""Strict, scalable even-odd reconstruction of source section contours."""
import numpy as np,shapely
from shapely.geometry import Polygon,GeometryCollection,LineString,box

def section_from_loops(loops,expected_area=None,allow_noded_contours=False):
    rings=[];paths=[];invalid=[];zero=0
    for loop in loops:
        if len(loop)<3:zero+=1;continue
        p=Polygon(loop)
        points=np.asarray(loop,dtype=float)
        if p.area==0 and np.linalg.matrix_rank(points-points[0])<2:zero+=1;continue
        if not p.is_valid:
            if not allow_noded_contours or expected_area is None:raise ValueError('Invalid source section contour; native-area proof required for source-segment arrangement')
            invalid.append(len(rings))
        rings.append(p);paths.append(points)
    if not rings:shape=GeometryCollection()
    else:
        tree=shapely.STRtree([box(*p.bounds) for p in rings] if invalid else rings)
        network=shapely.union_all([p.exterior for p in rings])
        cells=shapely.polygonize(shapely.get_parts(network));selected=[]
        for cell in cells.geoms:
            point=cell.representative_point()
            if not invalid:inside=len(tree.query(point,predicate='within'))%2
            else:
                inside=False;x,y=point.x,point.y
                for candidate in tree.query(point):
                    p=paths[int(candidate)];a=p;b=np.roll(p,-1,axis=0)
                    mask=(a[:,1]>y)!=(b[:,1]>y);aa=a[mask];bb=b[mask]
                    hits=aa[:,0]+(y-aa[:,1])*(bb[:,0]-aa[:,0])/(bb[:,1]-aa[:,1])
                    inside^=bool(np.count_nonzero(hits>x)%2)
            if inside:selected.append(cell)
        shape=shapely.union_all(selected)
    if not shape.is_valid:raise ValueError('Invalid reconstructed source section')
    error=abs(shape.area-expected_area) if expected_area is not None else None
    if error is not None and error>max(1e-4,abs(expected_area)*1e-9):raise ValueError('Source section area differs from native CrossSection: '+str(error))
    return shape,{'method':'native-area-verified-original-segment-arrangement-v3','contours':len(rings),'zero_area_contours':zero,'noded_touching_contours':invalid,
        'expected_native_area_mm2':expected_area,'reconstructed_area_mm2':float(shape.area),
        'area_error_mm2':error,'new_or_guessed_contours':0,'invalid_contours_automatically_repaired':False}
