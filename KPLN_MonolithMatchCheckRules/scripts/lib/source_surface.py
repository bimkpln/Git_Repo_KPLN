"""NWD source-surface correspondence and conservative candidate construction.

Requires numpy/scipy/trimesh/manifold3d/shapely. This is not the matching engine.
Every bounded change is exposed to downstream measurement uncertainty checks.
"""
import hashlib,json
from collections import defaultdict
import numpy as np
import trimesh,manifold3d as md
from scipy.spatial import cKDTree
from shapely.geometry import Polygon
from shapely.ops import unary_union,polygonize
import shapely
from topology_edges import split_contact_edges,split_t_seams,TopologyError,unit,sub,triangle_hash

def boundary_vertices(v,f,repeated=False):
    m=trimesh.Trimesh(vertices=v,faces=f,process=False)
    counts=np.bincount(m.edges_unique_inverse)
    return np.unique(m.edges_unique[counts!=2 if repeated else counts==1])

def sew_vertices(v,f,tolerance):
    """Candidate correspondence only; no unrestricted spatial vertex merge."""
    v=np.asarray(v);f=np.asarray(f);bv=boundary_vertices(v,f)
    pairs=sorted(cKDTree(v[bv]).query_pairs(tolerance)) if len(bv) else []
    parent=np.arange(len(v))
    def root(a):
        while parent[a]!=a:parent[a]=parent[parent[a]];a=parent[a]
        return a
    for a,b in pairs:parent[root(int(bv[b]))]=root(int(bv[a]))
    target=np.array([root(i) for i in range(len(v))]);mapped=v[target]
    displacements=np.linalg.norm(v-mapped,axis=1)
    if displacements.max(initial=0)>tolerance:raise TopologyError('Transitive correspondence exceeds bound')
    newfaces=target[f];keep=np.array([len(set(face))==3 for face in newfaces])
    source_norm=np.cross(v[f[:,1]]-v[f[:,0]],v[f[:,2]]-v[f[:,0]])
    mapped_norm=np.cross(v[newfaces[:,1]]-v[newfaces[:,0]],v[newfaces[:,2]]-v[newfaces[:,0]])
    if np.any(np.einsum('ij,ij->i',source_norm[keep],mapped_norm[keep])<=0):
        raise TopologyError('Vertex correspondence reverses source face')
    return v,newfaces[keep],{
        'method':'boundary-vertex-affine-correspondence-v1','tolerance_mm':tolerance,
        'maximum_displacement_mm':float(displacements.max(initial=0)),
        'source_vertex_to_output_vertex':target.tolist(),
        'collapsed_source_faces':np.flatnonzero(~keep).tolist(),
        'collapsed_source_area_mm2':float(np.linalg.norm(source_norm[~keep],axis=1).sum()/2),
        'source_triangle_count':len(f),'output_triangle_count':int(keep.sum()),
        'new_material_patches':0,'requires_downstream_uncertainty_accounting':True}

def subdivide_seams(v,f,tolerance,repeated=False):
    v=np.asarray(v);f=np.asarray(f)
    bv=boundary_vertices(v,f);tree=cKDTree(v[bv]) if len(bv) else None
    return split_t_seams(v,f,tolerance,(lambda c,r:bv[tree.query_ball_point(c,r)]) if tree else None,
        include_repeated_edges=repeated,reference_evidence=True)

def verify_subdivision(source_v,source_f,out_v,out_f,certificate):
    """Independently prove source-face support, coverage and displacement.

    A projected reference partition must cover every source triangle exactly,
    without overlaps. Affine vertex bounds extend to every point of each face.
    """
    owners=np.asarray(certificate['source_face_for_output_triangle']);refs=np.asarray(certificate['reference_triangle_coordinates'])
    actual=np.asarray(out_v)[np.asarray(out_f)];source=np.asarray(source_v)[np.asarray(source_f)]
    if len(refs)!=len(actual) or len(owners)!=len(actual):raise TopologyError('Missing source correspondence')
    maximum=float(np.linalg.norm(actual-refs,axis=2).max(initial=0))
    if maximum>certificate['seam_tolerance_mm']+1e-8:raise TopologyError('Actual displacement exceeds seam bound')
    max_coverage=0.;max_overlap=0.;max_plane=0.
    # Only split faces need polygon arrangement checks. Unchanged faces must
    # retain the exact ordered coordinates of their source triangle.
    grouped=defaultdict(list)
    for i,owner in enumerate(owners):grouped[int(owner)].append(i)
    if set(grouped)!=set(range(len(source))):raise TopologyError('Source face lost')
    for owner,indices in grouped.items():
        p=source[owner];rr=refs[indices]
        if len(indices)==1 and np.array_equal(rr[0],p):continue
        u=p[1]-p[0];u=u/np.linalg.norm(u);normal=np.cross(p[1]-p[0],p[2]-p[0]);normal/=np.linalg.norm(normal);w=np.cross(normal,u)
        plane=float(np.max(np.abs((rr-p[0])@normal)));max_plane=max(max_plane,plane)
        if plane>1e-7:raise TopologyError('Reference left original source plane')
        original=Polygon(np.column_stack(((p-p[0])@u,(p-p[0])@w)))
        polygons=[Polygon(np.column_stack(((r-p[0])@u,(r-p[0])@w))) for r in rr]
        combined=unary_union(polygons)
        coverage=original.symmetric_difference(combined).area
        overlap=abs(sum(x.area for x in polygons)-combined.area)
        max_coverage=max(max_coverage,coverage);max_overlap=max(max_overlap,overlap)
        if max(coverage,overlap)>max(1e-5,original.area*1e-9):raise TopologyError('Source coverage mismatch')
    return {'passed':True,'all_source_faces_covered':True,'maximum_affine_displacement_mm':maximum,
        'maximum_reference_plane_error_mm':max_plane,'maximum_per_face_coverage_error_mm2':max_coverage,
        'maximum_per_face_overlap_error_mm2':max_overlap}

def reduce_opposite_planes(v,faces,tolerance=.0001):
    """Retessellate signed source coverage, canceling internal double faces.

    Cells keep the full signed winding multiplicity. No polygon absent from
    the source coverage can be introduced. Projection error is explicit.
    """
    tri=v[faces];norm=np.cross(tri[:,1]-tri[:,0],tri[:,2]-tri[:,0]);length=np.linalg.norm(norm,axis=1);groups=defaultdict(list)
    if np.any(length<1e-12):raise TopologyError('Zero area source triangle')
    for i,n in enumerate(norm):
        n=n/length[i];sign=1 if n[np.argmax(np.abs(n))]>0 else -1;n=n*sign
        groups[tuple(np.round(n,7))+(round(float(n@tri[i,0])/tolerance),)].append((i,sign))
    output=v.tolist();newfaces=[];receipts=[];max_projection=0.
    for members in groups.values():
        if len({s for _,s in members})==1:
            newfaces.extend(faces[i].tolist() for i,_ in members);continue
        first,sign=members[0];n=norm[first]/length[first]*sign;origin=tri[first,0];u=np.asarray(unit(sub(tri[first,1],origin)));w=np.cross(n,u)
        polygons=[Polygon(np.column_stack(((tri[i]-origin)@u,(tri[i]-origin)@w))) for i,_ in members]
        if any(not p.is_valid or p.area==0 for p in polygons):raise TopologyError('Invalid projected source polygon')
        displacement=max(float(np.max(np.abs((tri[i]-origin)@n))) for i,_ in members)
        if displacement>tolerance:raise TopologyError('Source projection exceeds bound')
        plus=unary_union([p for p,(_,s) in zip(polygons,members) if s==1]);minus=unary_union([p for p,(_,s) in zip(polygons,members) if s==-1]);overlap=plus.intersection(minus).area
        if overlap<1e-8:
            newfaces.extend(faces[i].tolist() for i,_ in members);continue
        max_projection=max(max_projection,displacement)
        cells=list(polygonize(unary_union([p.boundary for p in polygons])))
        area_before=sum(p.area*s for p,(_,s) in zip(polygons,members));area_after=0.
        for cell in cells:
            point=cell.representative_point();winding=sum(s for p,(_,s) in zip(polygons,members) if p.covers(point))
            if not winding:continue
            triangles=shapely.constrained_delaunay_triangles(cell)
            if abs(sum(t.area for t in triangles.geoms)-cell.area)>max(1e-7,cell.area*1e-10):raise TopologyError('Triangulation loses signed source coverage')
            for t in triangles.geoms:
                xy=np.asarray(t.exterior.coords)[:3];p=origin+xy[:,0,None]*u+xy[:,1,None]*w
                if float(np.cross(p[1]-p[0],p[2]-p[0])@n)*winding<0:p=p[[0,2,1]]
                for _ in range(abs(winding)):
                    start=len(output);output.extend(p.tolist());newfaces.append([start,start+1,start+2])
            area_after+=cell.area*winding
        error=abs(area_after-area_before)
        if error>max(1e-5,sum(p.area for p in polygons)*1e-9):raise TopologyError('Signed source coverage changed')
        receipts.append({'source_faces':[i for i,_ in members],'opposite_overlap_area_mm2':overlap,'signed_area_balance_error_mm2':error})
    arr=np.asarray(output);vv,ii=np.unique(np.round(arr,7),axis=0,return_inverse=True)
    return vv,ii[np.asarray(newfaces)],{'method':'signed-source-plane-arrangement-v1','planes':receipts,
        'maximum_projection_mm':max_projection,'coordinate_rounding_bound_mm':np.sqrt(3)*.5e-7,'new_caps':0}

def native_candidate(v,f):
    v,f,cert=split_contact_edges(v,f) # Ambiguous zero-angle sectors are rejected.
    v=np.asarray(v,dtype='float64');f=np.asarray(f,dtype='uint64')
    center=(v.min(0)+v.max(0))/2
    native=md.Manifold(md.Mesh64(vert_properties=np.ascontiguousarray(v-center),tri_verts=f)).translate(center.tolist())
    if native.status()!=md.Error.NoError or native.volume()<=0:raise TopologyError('Native candidate rejected')
    return v,f,native,cert

def recover(v,f,max_seam_mm=.35):
    """Return a candidate plus fully disclosed correspondence chain.

    No result of this function alone is a complete geometry acceptance.
    """
    v=np.asarray(v,dtype=float);f=np.asarray(f,dtype=np.int64)
    if not np.all(np.isfinite(v)):raise TopologyError('Nonfinite source coordinates')
    attempts=[]
    def finish(vv,ff,steps,bound):
        bound=float(bound)
        rv,rf,native,cc=native_candidate(vv,ff)
        return rv,rf,native,{'source_sha256':triangle_hash(v,f),'source_faces':len(f),
            'steps':steps+[{'contact':cc}],'maximum_surface_correspondence_bound_mm':bound,
            'requires_csg_acceptance':True,'requires_measurement_uncertainty_accounting':bound>0}
    try:return finish(v,f,[],0.)
    except TopologyError as e:attempts.append(str(e))
    for weld_tol in [0,.0001,.005,.05,.15]:
        try:
            vv,ff,wc=sew_vertices(v,f,weld_tol) if weld_tol else (v,f,None)
            steps=[{'vertex_correspondence':wc}] if wc else [];bound=wc['maximum_displacement_mm'] if wc else 0.
            for seam_tol in [.0025,.01,.05]:
                try:
                    sv,sf,sc=subdivide_seams(vv,ff,seam_tol,repeated=True)
                    check=verify_subdivision(vv,ff,sv,sf,sc)
                    sc.pop('reference_triangle_coordinates');seamsteps=steps+[{'seam':sc,'independent_source_coverage':check}]
                    try:return finish(sv,sf,seamsteps,bound+check['maximum_affine_displacement_mm'])
                    except TopologyError:pass
                    # Only remaining boundary vertices participate in a second
                    # subdivision. This cannot pull unrelated nearby vertices.
                    if seam_tol==.01 and max_seam_mm>seam_tol:
                        try:
                            wv,wf,ws=subdivide_seams(sv,sf,max_seam_mm,repeated=True)
                            check2=verify_subdivision(sv,sf,wv,wf,ws);ws.pop('reference_triangle_coordinates')
                            return finish(wv,wf,seamsteps+[{'seam':ws,'independent_source_coverage':check2}],bound+check['maximum_affine_displacement_mm']+check2['maximum_affine_displacement_mm'])
                        except TopologyError:pass
                except TopologyError as e:attempts.append(str(e))
        except TopologyError as e:attempts.append(str(e))
    # Opposing coplanar faces describe an internal interface; signed coverage,
    # not triangle identity, is the invariant for this operation.
    pv,pf,pc=reduce_opposite_planes(v,f)
    if pc['planes']:
        vv,ff,wc=sew_vertices(pv,pf,.005)
        sv,sf,sc=subdivide_seams(vv,ff,.005,repeated=True);check=verify_subdivision(vv,ff,sv,sf,sc);sc.pop('reference_triangle_coordinates')
        return finish(sv,sf,[{'signed_plane_arrangement':pc},{'vertex_correspondence':wc},{'seam':sc,'independent_source_coverage':check}],pc['maximum_projection_mm']+pc['coordinate_rounding_bound_mm']+wc['maximum_displacement_mm']+check['maximum_affine_displacement_mm'])
    raise TopologyError('Source surface remains unresolved: '+','.join(sorted(set(attempts))))

def verify_native(v,f,native):
    """Independent signed volume plus live CSG partition/balance checks."""
    v=np.asarray(v);f=np.asarray(f);m=trimesh.Trimesh(vertices=v,faces=f,process=False)
    if not m.is_watertight or not m.is_winding_consistent:raise TopologyError('Indexed source surface invalid')
    tri=v[f]-v.mean(0);signed=float(np.einsum('ij,ij->i',tri[:,0],np.cross(tri[:,1],tri[:,2])).sum()/6)
    volume=float(native.volume());allowed=max(.01,abs(volume)*1e-9);errors={}
    errors['signed_volume']=abs(signed-volume)
    errors['self_difference']=abs((native-native).volume())
    for axis in range(3):
        normal=[0.,0.,0.];normal[axis]=1.
        a,b=native.split_by_plane(normal,float((v[:,axis].min()+v[:,axis].max())/2))
        if any(x.status()!=md.Error.NoError for x in [a,b]):raise TopologyError('Live CSG split failed')
        errors['partition_'+str(axis)]=abs(a.volume()+b.volume()-volume)
        errors['partition_overlap_'+str(axis)]=abs((a^b).volume())
        errors['reunion_'+str(axis)]=abs((a+b).volume()-volume)
    if any(x>allowed for x in errors.values()):raise TopologyError('Native CSG volume balance failed: '+str(errors))
    return {'passed':True,'native_volume_mm3':volume,'independent_signed_volume_mm3':signed,
        'numerical_volume_budget_mm3':allowed,'volume_balance_errors_mm3':errors,
        'live_csg_partition_axes':3,'all_source_faces_watertight':True}
