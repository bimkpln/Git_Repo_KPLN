"""Optional original-local NWD coordinates, before far-world float additions.

This checks source provenance and literal callback recomposition only. It
does not certify topology, volumes, admission roles or a completed comparison.
"""
from fractions import Fraction
from pathlib import Path
import hashlib,math,re,struct,sys

NAME='nwd_local_triangles.bin'


def filehash(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda:stream.read(4*1024*1024),b''):h.update(block)
    return h.hexdigest()


def matrix(value):
    if len(value)!=16 or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x) for x in value):
        raise ValueError('Finite literal source matrix required')
    if any(value[k]!=0 for k in [3,7,11]) or value[15]!=1:
        raise ValueError('A literal affine source matrix is required')
    return value


def callback_point(point,transform):
    return tuple(point[0]*transform[i]+point[1]*transform[i+4]+point[2]*transform[i+8]+transform[i+12] for i in range(3))


def exact_affine_point(point,transform):
    transform=matrix(transform)
    if len(point)!=3 or any(isinstance(x,bool) or not math.isfinite(x) for x in point):raise ValueError('Three finite local coordinates required')
    p=tuple(Fraction(x) for x in point);m=tuple(Fraction(x) for x in transform)
    return tuple(sum(p[j]*m[i+4*j] for j in range(3))+m[12+i] for i in range(3))


def validate_local_reference(directory,manifest,fragments):
    root=Path(directory);local=root/NAME
    name=manifest.get('local_triangle_file');sha=manifest.get('local_triangle_sha256')
    if name is None and sha is None and not local.exists():return None
    if name!=NAME or not isinstance(sha,str) or not re.fullmatch('[0-9a-f]{64}',sha):
        raise ValueError('Fixed local-coordinate file name and SHA256 required')
    world=root/'nwd_triangles.bin'
    if not local.is_file() or local.stat().st_size!=world.stat().st_size or filehash(local)!=sha:
        raise ValueError('Truncated, changed or missing local-coordinate sidecar')
    cursor=0;vertices=0;maximum=0.
    with local.open('rb') as a,world.open('rb') as b:
        for f in fragments:
            offset=f['triangle_offset'];count=f['triangle_count']
            if isinstance(count,bool) or not isinstance(count,int) or count<0 or offset!=cursor:
                raise ValueError('Non-contiguous local-coordinate fragment range')
            M=matrix(f['local_to_world']);remaining=count*72
            while remaining:
                size=min(remaining,72*4096);left=a.read(size);right=b.read(size)
                if len(left)!=size or len(right)!=size:raise ValueError('Truncated source-coordinate range')
                for p,q in zip(struct.iter_unpack('<3d',left),struct.iter_unpack('<3d',right)):
                    if not all(math.isfinite(x) for x in p+q):raise ValueError('Non-finite source coordinate')
                    expected=callback_point(p,M)
                    if expected!=q:raise ValueError('Local coordinates do not reproduce literal world callback')
                    for i in range(3):
                        absolute=sum(abs(p[j]*M[i+4*j]) for j in range(3))+abs(M[i+12])
                        bound=absolute*sys.float_info.epsilon*8+math.ulp(0.)*8
                        if not math.isfinite(bound):raise ValueError('Unbounded source affine arithmetic')
                        maximum=max(maximum,bound)
                    vertices+=1
                remaining-=size;cursor+=size
    if cursor!=local.stat().st_size:raise ValueError('Trailing or unaccounted source-coordinate bytes')
    return {'passed':True,'file':NAME,'sha256':sha,'vertices':vertices,'triangles':vertices//3,
        'literal_callback_recomposition_matches_all_vertices':True,
        'maximum_callback_arithmetic_bound_input_units':maximum,
        'exact_affine_reference_available':True,'coordinate_rounding_or_welding':False,
        'scope':'source-coordinate reference only; topology and solid acceptance remain required',
        'complete_geometry_check':False,'release_allowed':False}
