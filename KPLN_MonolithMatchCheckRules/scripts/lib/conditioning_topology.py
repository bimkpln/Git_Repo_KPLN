"""Reject conditioning which changes connected shells or tunnel topology."""
from source_surface import TopologyError
def topology_signature(native):
    return sorted((int(part.genus()),1 if part.volume()>0 else -1 if part.volume()<0 else 0) for part in native.decompose())
def require_same_topology(reference,candidate):
    a=topology_signature(reference);b=topology_signature(candidate)
    if a!=b:raise TopologyError('Conditioning changes source topology: '+str(a)+' -> '+str(b))
    return {'passed':True,'original_components_and_genera':a,'conditioned_components_and_genera':b,
        'invariants_alone_do_not_replace_surface_or_volume_checks':True}
