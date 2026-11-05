"""Canonical triangulation of existing signed planar source skin.

No hulls, caps or guessed planes. This conditioning is independently checked
before any source surface can participate in the full-model CSG.
"""
from source_surface import *

def canonical_skin(v,f,tolerance_mm=.0001,sew_mm=.005):
    v=np.asarray(v,dtype=float);f=np.asarray(f,dtype=np.int64);tri=v[f]
    norm=np.cross(tri[:,1]-tri[:,0],tri[:,2]-tri[:,0]);length=np.linalg.norm(norm,axis=1);groups=defaultdict(list)
    for i,n in enumerate(norm):
        if length[i]<=0:raise TopologyError('Degenerate source skin')
        n=n/length[i];sign=1 if n[np.argmax(abs(n))]>0 else -1;n=n*sign
        groups[tuple(np.round(n,7))+(round(float(n@tri[i,0])/tolerance_mm),)].append((i,sign))
    vv=[];ff=[];proof=[];maximum_plane_error=0.
    for members in groups.values():
        i,sign=members[0];normal=norm[i]/length[i]*sign;origin=tri[i,0];u=tri[i,1]-origin;u=u/np.linalg.norm(u);w=np.cross(normal,u)
        pp=[Polygon(np.c_[(tri[j]-origin)@u,(tri[j]-origin)@w]) for j,_ in members]
        if any(not p.is_valid or p.area<=0 for p in pp):raise TopologyError('Invalid source skin polygon')
        maxdist=max(float(np.abs((tri[j]-origin)@normal).max()) for j,_ in members)
        if maxdist>tolerance_mm:raise TopologyError('Source plane spread exceeds tolerance')
        maximum_plane_error=max(maximum_plane_error,maxdist)
        cells=list(polygonize(unary_union([p.boundary for p in pp])));signed=defaultdict(list)
        for cell in cells:
            p=cell.representative_point();wind=sum(s for poly,(_,s) in zip(pp,members) if poly.covers(p))
            if wind:signed[wind].append(cell)
        output_area=0.;max_coverage=0.
        for wind,polys in signed.items():
            full=unary_union(polys)
            regions=[full] if full.geom_type=='Polygon' else list(full.geoms)
            for region in regions:
                ts=shapely.constrained_delaunay_triangles(region)
                triangles=list(ts.geoms);coverage=region.symmetric_difference(unary_union(triangles)).area
                overlap=abs(sum(t.area for t in triangles)-region.area)
                max_coverage=max(max_coverage,coverage,overlap)
                if max(coverage,overlap)>max(1e-5,region.area*1e-9):raise TopologyError('Canonical skin triangulation loses coverage')
                for t in triangles:
                    xy=np.asarray(t.exterior.coords)[:3];p=origin+xy[:,0,None]*u+xy[:,1,None]*w
                    if np.dot(np.cross(p[1]-p[0],p[2]-p[0]),normal)*wind<0:p=p[[0,2,1]]
                    for _ in range(abs(wind)):
                        k=len(vv);vv.extend(p.tolist());ff.append([k,k+1,k+2])
                output_area+=region.area*wind
        error=abs(output_area-sum(p.area*s for p,(_,s) in zip(pp,members)))
        if error>max(1e-5,sum(p.area for p in pp)*1e-9):raise TopologyError('Signed source skin coverage changed')
        proof.append({'source_face_indices':[j for j,_ in members],'source_faces':len(members),'signed_area_error_mm2':error,'plane_displacement_mm':maxdist,'triangulation_coverage_error_mm2':max_coverage})
    v0,ii=np.unique(np.round(vv,7),axis=0,return_inverse=True);f0=ii[np.asarray(ff)]
    v1,f1,weld=sew_vertices(v0,f0,sew_mm)
    v2,f2,seam=subdivide_seams(v1,f1,sew_mm,True);coverage=verify_subdivision(v1,f1,v2,f2,seam);seam.pop('reference_triangle_coordinates')
    v3,f3,native,contacts=native_candidate(v2,f2);native_check=verify_native(v3,f3,native)
    bound=maximum_plane_error+float(np.sqrt(3)*.5e-7)+weld['maximum_displacement_mm']+coverage['maximum_affine_displacement_mm']
    certificate={'source_triangle_sha256':triangle_hash(v,f),'method':'canonical-signed-source-planar-skin-v1',
        'source_faces':len(f),'output_faces':len(f3),'planes':proof,'vertex_correspondence':weld,
        'seam_correspondence':seam,'independent_source_coverage':coverage,'contact_identity':contacts,
        'native_acceptance':native_check,'maximum_additional_surface_correspondence_bound_mm':bound,
        'new_or_guessed_caps':0,'requires_downstream_uncertainty_accounting':True}
    return v3,f3,native,certificate
