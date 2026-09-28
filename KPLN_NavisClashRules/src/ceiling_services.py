"""AR ceilings: longitudinal services and verified transverse duct sections."""

import math

DISCIPLINES = {"OV", "VK", "PT", "EOM", "SS", "ОВ", "ВК", "ПТ", "ЭОМ", "СС"}


def unit_vector(value):
    if isinstance(value, dict):
        value = [value.get(axis) for axis in ("X", "Y", "Z")]
    if not isinstance(value, (list, tuple)) or len(value) != 3:
        return None
    try:
        if any(isinstance(component, bool) for component in value):
            return None
        vector = [float(component) for component in value]
        if not all(math.isfinite(component) for component in vector):
            return None
        length = math.hypot(*vector)
        return [component / length for component in vector] if length else None
    except (TypeError, ValueError):
        return None


def classify_ceiling(row, discipline):
    if discipline not in DISCIPLINES:
        raise ValueError("Ceiling services require OV, VK, PT, EOM or SS discipline")
    segments = {s.strip().casefold() for key in ("Item1Path", "Item2Path")
                for s in (row.get(key) or "").split(" / ")}
    duct_segments = {"воздуховоды", "ducts", "материалы изоляции воздуховодов",
                     "duct insulations", "duct insulation", "изоляция воздуховода"}
    duct_classes = {"воздуховод", "изоляция воздуховода"}
    is_duct = bool(segments & duct_segments) or bool(duct_classes & {row.get("Class1"), row.get("Class2")})
    match = row.get("DuctSectionFootprintMatch")
    source = row.get("DuctSectionCheckSource")
    expected_source = "object-section-properties+duct-ceiling-mesh-section"
    if is_duct or source == expected_source:
        if source != expected_source or not isinstance(match, bool):
            return "Uncertain", "missing:duct_section_footprint"
        if match:
            return "Approved", "transverse_duct_section_matches_ceiling_footprint"
    # Explicit normalized geometry, never an axis guessed from AABB dimensions.
    axis = unit_vector(row.get("MepAxis"))
    normal = unit_vector(row.get("CeilingNormal"))
    missing = [key for key, vector in (("MepAxis", axis), ("CeilingNormal", normal)) if vector is None]
    if missing:
        return "Uncertain", "missing:" + ",".join(missing)
    normal_component = min(1.0, abs(sum(a * b for a, b in zip(axis, normal))))
    plane_component = math.sqrt(max(0.0, 1.0 - normal_component ** 2))
    if plane_component > normal_component and not math.isclose(plane_component, normal_component, abs_tol=1e-12):
        return "Approved", "longitudinal_service_in_ar_ceiling"
    return "Uncertain", "transverse_or_boundary_outside_longitudinal_scope"
