"""Independent empty-volume predicate for a closed convex box.

Boundary separation alone is insufficient: an entire material component can
lie inside the box, or the box can lie entirely inside material.
"""
import numpy as np

def empty_box_certificate(boundary_triangle_minima,source_vertices_inside,center_outside_all_sources,
                          uncertainty_mm,numerical_guard_mm,all_source_faces_vertices_and_membership_covered):
    values=np.asarray(boundary_triangle_minima,float)
    if values.shape!=(12,) or not np.isfinite(values).all() or np.any(values<0):raise ValueError('Twelve finite nonnegative boundary triangle minima are required')
    if uncertainty_mm<0 or numerical_guard_mm<=0:raise ValueError('Valid explicit uncertainty and positive numeric guard are required')
    separated=bool(np.all(values>uncertainty_mm+numerical_guard_mm))
    passed=bool(separated and not source_vertices_inside and center_outside_all_sources and all_source_faces_vertices_and_membership_covered)
    return {'passed':passed,'method':'closed-convex-box-boundary-separation-vertex-exclusion-and-independent-membership-v1',
        'minimum_boundary_distance_mm':float(values.min()),'source_vertices_inside':bool(source_vertices_inside),
        'center_outside_all_sources':bool(center_outside_all_sources),'uncertainty_mm':float(uncertainty_mm),
        'numerical_guard_mm':float(numerical_guard_mm),'all_sources_covered':bool(all_source_faces_vertices_and_membership_covered),
        'mathematical_test_domain_not_model_material':True}
