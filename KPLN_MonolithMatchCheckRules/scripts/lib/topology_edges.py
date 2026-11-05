"""Separate touching oriented surface sheets without adding any geometry."""
import math,hashlib,json
from collections import defaultdict,Counter

class TopologyError(ValueError):pass

def sub(a,b):return tuple(x-y for x,y in zip(a,b))
def dot(a,b):return sum(x*y for x,y in zip(a,b))
def cross(a,b):return (a[1]*b[2]-a[2]*b[1],a[2]*b[0]-a[0]*b[2],a[0]*b[1]-a[1]*b[0])
def unit(a):
    length=math.sqrt(dot(a,a))
    if not length:raise TopologyError('Zero length source edge')
    return tuple(x/length for x in a)

def triangle_hash(vertices,faces):
    h=hashlib.sha256()
    for face in faces:
        h.update(json.dumps([[float(x) for x in vertices[v]] for v in face],separators=(',',':'),allow_nan=False).encode());h.update(b'\n')
    return h.hexdigest()

def split_contact_edges(vertices,faces,allow_zero_sectors=False):
    """Pair each outward-oriented material sector around a repeated edge.

    Every triangle and its coordinates/winding are preserved. Only vertex
    identities change. Open edges and ambiguous angular sectors are rejected.
    A successful return still requires native CSG/volume acceptance by caller.
    """
    vertices=[tuple(map(float,v)) for v in vertices];faces=[tuple(map(int,f)) for f in faces]
    edges=defaultdict(list)
    for f,face in enumerate(faces):
        if len(face)!=3 or len(set(face))!=3:raise TopologyError('Degenerate source face')
        for i in range(3):
            a,b,c=face[i],face[(i+1)%3],face[(i+2)%3]
            edges[tuple(sorted((a,b)))].append((f,i,a<b,c))
    parent=list(range(len(faces)*3))
    def root(i):
        while parent[i]!=i:parent[i]=parent[parent[i]];i=parent[i]
        return i
    def join(a,b):
        a,b=root(a),root(b)
        if a!=b:parent[b]=a
    contact_count=0;zero_sectors=0;paired_edge_ids={};edge_pairs=[]
    for (a,b),members in edges.items():
        if len(members)==1:raise TopologyError('Open source edge')
        if len(members)%2 or sum(m[2] for m in members)*2!=len(members):
            raise TopologyError('Unbalanced source face winding')
        if len(members)==2:pairs=[members]
        else:
            contact_count+=1;e=unit(sub(vertices[b],vertices[a]))
            axis=min([(1.,0.,0.),(0.,1.,0.),(0.,0.,1.)],key=lambda x:abs(dot(x,e)))
            u=unit(cross(e,axis));v=cross(e,u)
            ordered=sorted([(math.atan2(dot(sub(vertices[m[3]],vertices[a]),v),
                                      dot(sub(vertices[m[3]],vertices[a]),u)),m) for m in members],key=lambda x:x[0])
            pairs=[]
            for index,(angle,m) in enumerate(ordered):
                previous_angle,previous=ordered[index-1]
                separation=(angle-previous_angle)%(2*math.pi)
                if separation<1e-10:
                    zero_sectors+=1
                    if not allow_zero_sectors:raise TopologyError('Coincident source sectors require separate proof')
                if m[2]:
                    if previous[2]:raise TopologyError('Overlapping or inconsistent material sectors')
                    pairs.append([m,previous])
        for first,second in pairs:
            f,i,forward,_=first;g,j,backward,_=second
            if forward==backward:raise TopologyError('Same-direction face pairing')
            # Edge endpoints are the same coordinates, in opposite directions.
            join(f*3+i,g*3+(j+1)%3);join(f*3+(i+1)%3,g*3+j)
            edge_id=len(edge_pairs);edge_pairs.append((a,b))
            paired_edge_ids[(f,i)]=edge_id;paired_edge_ids[(g,j)]=edge_id
    output_vertices=[];output_faces=[];indices={}
    for f,face in enumerate(faces):
        result=[]
        for i,source in enumerate(face):
            identity=root(f*3+i)
            if identity not in indices:indices[identity]=len(output_vertices);output_vertices.append(vertices[source])
            result.append(indices[identity])
        output_faces.append(result)
    degree=Counter(tuple(sorted((f[i],f[(i+1)%3]))) for f in output_faces for i in range(3))
    subdivided=False
    if any(n!=2 for n in degree.values()):
        # A valid paired-edge surface may contain parallel topological edges
        # with identical endpoints. Barycentric subdivision distinguishes
        # those identities while partitioning every original triangle exactly.
        subdivided=True;corner_faces=output_faces;output_faces=[]
        midpoint_ids=[]
        for a,b in edge_pairs:
            midpoint_ids.append(len(output_vertices))
            output_vertices.append(tuple((x+y)/2 for x,y in zip(vertices[a],vertices[b])))
        for f,face in enumerate(faces):
            center_id=len(output_vertices)
            output_vertices.append(tuple(sum(vertices[k][j] for k in face)/3 for j in range(3)))
            for i in range(3):
                a=corner_faces[f][i];b=corner_faces[f][(i+1)%3];mid=midpoint_ids[paired_edge_ids[(f,i)]]
                output_faces.extend([[a,mid,center_id],[mid,b,center_id]])
        degree=Counter(tuple(sorted((f[i],f[(i+1)%3]))) for f in output_faces for i in range(3))
        if any(n!=2 for n in degree.values()):raise TopologyError('Barycentric surface is nonmanifold')
    before=triangle_hash(vertices,faces);after=triangle_hash(output_vertices,output_faces)
    if not subdivided and before!=after:raise AssertionError('Source triangle geometry changed')
    return output_vertices,output_faces,{'method':'oriented-source-material-sectors-v1',
        'source_triangle_geometry_sha256':before,'output_triangle_geometry_sha256':after,
        'contact_edges_split':contact_count,'source_vertices':len(vertices),'output_vertices':len(output_vertices),
        'source_faces':len(faces),'output_faces':len(output_faces),
        'new_surface_faces':0,'new_coordinates':len(edge_pairs)+len(faces) if subdivided else 0,
        'source_triangle_geometry_unchanged':not subdivided,'exact_barycentric_partition':subdivided,
        'zero_angle_sector_boundaries':zero_sectors,'native_volume_acceptance_required':True}

def split_t_seams(vertices,faces,tolerance_mm=.005,query_vertices=None,include_repeated_edges=False,reference_evidence=False):
    """Subdivide an existing face at vertices of its mismatched boundary seam.

    No open loop is capped. Reference subdivisions partition the original
    triangle exactly. Reusing nearby existing seam vertices gives a rigorous
    affine displacement bound, recorded for both directions of correspondence.
    """
    vertices=[tuple(map(float,v)) for v in vertices];faces=[tuple(map(int,f)) for f in faces]
    edges=defaultdict(list)
    for f,face in enumerate(faces):
        for i in range(3):edges[tuple(sorted((face[i],face[(i+1)%3])))].append((f,i))
    boundary={k:v[0] for k,v in edges.items() if len(v)==1 or (include_repeated_edges and len(v)!=2)}
    candidates=sorted({v for edge in boundary for v in edge})
    insertions={};max_displacement=0.;split_edges=0
    for (a,b),(f,i) in boundary.items():
        aa,bb=vertices[a],vertices[b];ab=sub(bb,aa);l2=dot(ab,ab)
        if not l2:raise TopologyError('Zero length seam edge')
        center=tuple((x+y)/2 for x,y in zip(aa,bb));radius=math.sqrt(l2)/2+tolerance_mm
        near=candidates if query_vertices is None else query_vertices(center,radius)
        members=[]
        for c in near:
            if c in faces[f]:continue
            p=vertices[c];t=dot(sub(p,aa),ab)/l2
            if t<=1e-10 or t>=1-1e-10:continue
            ref=tuple(x+t*d for x,d in zip(aa,ab));delta=sub(p,ref);distance=math.sqrt(dot(delta,delta))
            if distance<=tolerance_mm:
                members.append((t,c,ref,distance));max_displacement=max(max_displacement,distance)
        if members:
            members.sort()
            if any(y[0]-x[0]<1e-10 for x,y in zip(members,members[1:])):
                raise TopologyError('Ambiguous seam vertices at the same source-edge parameter')
            for owner,edge_index in edges[(a,b)]:
                insertions[(owner,edge_index)]=members if faces[owner][edge_index]==a else list(reversed(members))
            split_edges+=1
    output_vertices=list(vertices);output_faces=[];owners=[];references=[];reference_area=0.;original_area=0.
    for f,face in enumerate(faces):
        p,q,r=[vertices[v] for v in face];normal=cross(sub(q,p),sub(r,p));area=math.sqrt(dot(normal,normal))/2
        original_area+=area
        if not any((f,i) in insertions for i in range(3)):
            output_faces.append(list(face));owners.append(f);references.append([p,q,r]);reference_area+=area;continue
        polygon=[]
        for i in range(3):
            polygon.append((face[i],vertices[face[i]]))
            for _,c,ref,_ in insertions.get((f,i),[]):polygon.append((c,ref))
        centroid=tuple((a+b+c)/3 for a,b,c in zip(p,q,r));center_id=len(output_vertices);output_vertices.append(centroid)
        for j,(a,refa) in enumerate(polygon):
            b,refb=polygon[(j+1)%len(polygon)]
            actual_normal=cross(sub(output_vertices[b],output_vertices[a]),sub(centroid,output_vertices[a]))
            if dot(actual_normal,normal)<=0:raise TopologyError('Seam subdivision reverses a source surface')
            refnormal=cross(sub(refb,refa),sub(centroid,refa));reference_area+=math.sqrt(dot(refnormal,refnormal))/2
            output_faces.append([a,b,center_id]);owners.append(f)
            references.append([refa,refb,centroid])
    area_error=abs(reference_area-original_area)
    if area_error>max(1e-5,original_area*1e-10):raise TopologyError('Reference subdivision does not preserve source surface area')
    certificate={'method':'source-face-affine-seam-subdivision-v1',
        'split_source_edges':split_edges,'source_faces':len(faces),'output_faces':len(output_faces),
        'source_face_for_output_triangle':owners,'source_reference_area_mm2':original_area,
        'reference_area_balance_error_mm2':area_error,'maximum_bidirectional_surface_displacement_bound_mm':max_displacement,
        'seam_tolerance_mm':tolerance_mm,'new_caps':0,'new_material_patches':0,
        'new_vertices_only_on_original_source_faces':True,'source_vertices_moved':0,
        'native_volume_acceptance_required':True}
    if reference_evidence:certificate['reference_triangle_coordinates']=references
    return output_vertices,output_faces,certificate
