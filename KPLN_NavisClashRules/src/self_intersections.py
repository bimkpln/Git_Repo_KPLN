"""Classify trained pipe/duct self-intersection results."""

from __future__ import annotations

import math
from typing import Any


BARE_KIND = "голая труба"
INSULATED_KINDS = {"изоляция", "фитинг"}
SUPPORTED_KINDS = INSULATED_KINDS | {BARE_KIND}


def as_float(value: Any) -> float | None:
    try:
        if value is None or isinstance(value, bool):
            return None
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def classify_self_intersection(row: dict[str, Any]) -> tuple[str, str]:
    kind1 = str(row.get("Kind1") or "").strip().lower()
    kind2 = str(row.get("Kind2") or "").strip().lower()
    kinds = {kind1, kind2}

    if kind1 not in SUPPORTED_KINDS or kind2 not in SUPPORTED_KINDS:
        return "Uncertain", "missing:supported_self_intersection_kind"

    if kind1 == BARE_KIND and kind2 == BARE_KIND:
        return "Uncertain", "bare_body_pair_not_trained"

    # Fittings are treated as insulated because their insulation is not
    # exported to NWC. Contact with an ordinary bare body means that envelope
    # was already penetrated, regardless of the measured angle.
    if BARE_KIND in kinds and kinds & INSULATED_KINDS:
        return "Active", "insulated_envelope_reaches_bare_body"

    if row.get("GeomOk") is not True:
        return "Uncertain", "missing:reliable_self_intersection_geometry"

    perp_rel = as_float(row.get("PerpRel"))
    angle = as_float(row.get("Angle"))
    elong1 = as_float(row.get("Elong1"))
    elong2 = as_float(row.get("Elong2"))

    if perp_rel is not None and perp_rel <= 0.15:
        return "Active", "axis_crossing_perp_rel_at_or_below_0_15"

    elongations = [value for value in (elong1, elong2) if value is not None]
    if angle is not None and angle <= 5.0 and any(value >= 5.0 for value in elongations):
        return "Active", "parallel_extended_routes_overlap"

    missing = []
    if perp_rel is None:
        missing.append("PerpRel")
    if angle is None:
        missing.append("Angle")
    elif angle <= 5.0 and (elong1 is None or elong2 is None):
        missing.append("Elong1/Elong2")
    if missing:
        return "Uncertain", "missing:" + ",".join(missing)

    return "Approved", "transverse_or_local_insulated_contact"
