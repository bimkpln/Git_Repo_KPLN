"""Directed-contour reconstruction with exact planar-arrangement fallback."""
from fractions import Fraction
from native_section_geometry_nested import native_section_from_loops as nested_section
from rational_arrangement import positive_region

def native_section_from_loops(loops,expected_area,allow_noded_contours=False,representation_bound_mm=1e-7):
    loops=list(loops)
    if representation_bound_mm<0:raise ValueError('Negative representation bound')
    try:
        return nested_section(loops,expected_area,allow_noded_contours,representation_bound_mm)
    except ValueError as failure:
        if not allow_noded_contours:raise
        cycles=[[[Fraction(float(x)),Fraction(float(y))] for x,y in loop] for loop in loops]
        shape,arrangement=positive_region(cycles)
        if any(w<0 for w in arrangement['face_windings']):
            raise ValueError('Negative contour outside material') from failure
        error=abs(float(shape.area)-float(expected_area))
        if error>max(1e-4,abs(float(expected_area))*1e-9):
            raise ValueError('Exact directed-contour arrangement disagrees with native area') from failure
        if arrangement['maximum_coordinate_representation_error_mm']>representation_bound_mm:
            raise ValueError('Exact arrangement exceeds coordinate representation bound') from failure
        return shape,{'method':'native-directed-exact-arrangement-fallback-v1','nested_reconstruction_rejection':str(failure),
            'expected_native_area_mm2':float(expected_area),'reconstructed_area_mm2':float(shape.area),'area_error_mm2':error,
            'arrangement':arrangement,'representation_bound_mm':representation_bound_mm,
            'no_hole_size_filter':True,'new_or_guessed_contours':0,'model_geometry_modified':False}
