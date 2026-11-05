"""Constant union-material state in a fully audited connected box domain."""
import numpy as np

def box_material_certificate(boundary_triangle_minima,source_vertices_inside,center_membership,
                             uncertainty_mm,numerical_guard_mm,all_sources_covered):
    values=np.asarray(boundary_triangle_minima,float)
    if values.shape!=(12,) or not np.isfinite(values).all() or np.any(values<0):raise ValueError('Twelve finite boundary minima required')
    if center_membership not in ['outside','inside','uncertain']:raise ValueError('Explicit independent center membership required')
    if uncertainty_mm<0 or numerical_guard_mm<=0:raise ValueError('Explicit uncertainty and positive guard required')
    separated=bool(np.all(values>uncertainty_mm+numerical_guard_mm))
    state='unproved'
    if separated and not source_vertices_inside and all_sources_covered:
        state={'outside':'empty','inside':'occupied','uncertain':'unproved'}[center_membership]
    return {'state':state,'method':'constant-material-state-in-connected-domain-without-internal-source-boundary-v1',
        'minimum_boundary_distance_mm':float(values.min()),'source_vertices_inside':bool(source_vertices_inside),
        'independent_center_membership':center_membership,'uncertainty_mm':float(uncertainty_mm),
        'numerical_guard_mm':float(numerical_guard_mm),'all_sources_covered':bool(all_sources_covered),
        'whole_source_cut_not_disposed_by_core_certificate':True}
