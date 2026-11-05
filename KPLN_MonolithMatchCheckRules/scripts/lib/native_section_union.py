"""Native union of original material sections, independently reconstructed."""
import numpy as np
import shapely
import manifold3d as md
from rational_coverage import coverage as exact_coverage
from native_section_geometry import native_section_from_loops as section_from_loops
def assemble_sections(cross_sections,representation_bound_mm=0.):
    cross_sections=list(cross_sections);native=md.CrossSection.batch_boolean(cross_sections,md.OpType.Add)
    reconstruction_failure=None
    try:shape,reconstruction=section_from_loops(native.to_polygons(),float(native.area()),allow_noded_contours=True,representation_bound_mm=representation_bound_mm)
    except ValueError as e:reconstruction_failure=str(e)
    omitted=[float((part-native).area()) for part in cross_sections]
    parts=[section_from_loops(p.to_polygons(),float(p.area()),allow_noded_contours=True,representation_bound_mm=representation_bound_mm)[0] for p in cross_sections]
    if reconstruction_failure is not None:
        shape=shapely.union_all(parts)
        if not shape.is_valid:raise ValueError('Independent source section union is invalid')
        if abs(shape.area-float(native.area()))>max(1e-4,abs(float(native.area()))*1e-9):raise ValueError('Independent source union differs from native area')
        reconstruction={'method':'independent-valid-source-material-union','native_contour_rejection':reconstruction_failure,'area_error_mm2':abs(shape.area-float(native.area()))}
    if representation_bound_mm<0:raise ValueError('Negative representation bound')
    if any(a>0 for a in omitted):
        # A positive omission is never excused by its area. Prove the entire
        # original material lies inside a stated distance enclosure instead.
        if not representation_bound_mm:raise ValueError('Native section union omits original material')
        outside=[p.difference(shape.buffer(representation_bound_mm)).area for p in parts]
        if any(a>0 for a in outside):raise ValueError('Native section omission exceeds distance enclosure')
    if representation_bound_mm==0 and shape.difference(shapely.union_all(parts)).area>0:raise ValueError('Strict native union adds material to original source void')
    input_voids_checked=0;exact_false_void_proofs=[]
    part_tree=shapely.STRtree(parts)
    for part in parts:
        components=[part] if part.geom_type=='Polygon' else list(getattr(part,'geoms',[]))
        for component in components:
            if component.geom_type!='Polygon':continue
            for ring in component.interiors:
                free=shapely.Polygon(ring)
                for i in part_tree.query(free):free=free.difference(parts[int(i)])
                regions=[free] if free.geom_type=='Polygon' else list(getattr(free,'geoms',[]))
                for region in regions:
                    if region.geom_type!='Polygon' or region.area<=0:continue
                    input_voids_checked+=1
                    if region.difference(shape).area==0:
                        proof=exact_coverage(region,[parts[int(i)] for i in part_tree.query(region)])
                        if not proof['covered']:raise ValueError('A complete independent input void was filled, irrespective of its area')
                        exact_false_void_proofs.append(proof)
    precision=1e-8;other=shapely.union_all(shapely.set_precision(parts,precision),grid_size=precision)
    if representation_bound_mm:
        enclosure=shapely.union_all([p.buffer(representation_bound_mm) for p in parts])
        if shape.difference(enclosure).area>0:raise ValueError('Native section creates material beyond distance enclosure')
    disagreement=float(shape.symmetric_difference(other).area);budget=max(1e-4,sum(p.length for p in parts)*precision*4)
    if disagreement>budget:raise ValueError('Independent section union differs: '+str(disagreement))
    return native,shape,{'passed':True,'method':'original-source-native-cross-section-union-and-independent-area-parity',
        'source_parts':len(parts),'material_area_mm2':float(native.area()),'omitted_source_material_area_mm2':sum(omitted),
        'native_reconstruction':reconstruction,'independent_union_difference_mm2':disagreement,
        'independent_union_representation_grid_mm':precision,'representation_grid_changes_output_geometry':False,
        'hole_size_filter_applied':False,'independent_input_void_components_checked':input_voids_checked,'exact_false_void_proofs':exact_false_void_proofs,'maximum_section_boundary_error_mm':representation_bound_mm,
        'exact_original_material_inclusion':not any(a>0 for a in omitted),
        'distance_enclosure_is_not_building_tolerance':True,
        'absence_decisions_require_original_seed_and_boundary_band_audit':representation_bound_mm>0}
