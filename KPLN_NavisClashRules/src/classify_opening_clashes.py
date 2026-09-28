#!/usr/bin/env python3
"""Classify structure-vs-duct opening clashes from MCP result JSON.

Input must be a JSON array returned by get_clash_test_results with
include_metrics=True. Keep Bound in the input when bundle rules are needed.
The script contains no project names, clash numbers, or trained examples.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import Counter, defaultdict
from typing import Any

from provenance import provenance
from ceiling_services import DISCIPLINES, classify_ceiling
from slab_openings import classify_slabs


ROUNDING_NOISE_MM = 1.0
ROUND_DAMPER_SHALLOW_DEPTH_RATIO = 0.6


def status_bucket(value: Any) -> str:
    text = str(value or "").upper()
    if "APPROVED" in text or "RESOLVED" in text or "REVIEWED" in text:
        return "Approved"
    return "Active"


def group_key(value: Any) -> str | None:
    if not value:
        return None
    first = str(value).splitlines()[0].strip()
    return first if first else None


def path_text(row: dict[str, Any]) -> str:
    return f"{row.get('Item1Path') or ''} {row.get('Item2Path') or ''}".lower()


def is_round_damper(row: dict[str, Any]) -> bool:
    return row.get("PairClass") == "клапан+стена" and "кругл" in path_text(row)


def is_rectangular_damper(row: dict[str, Any]) -> bool:
    return row.get("PairClass") == "клапан+стена" and "прямоуголь" in path_text(row)


def is_round_elbow(row: dict[str, Any]) -> bool:
    text = path_text(row)
    return row.get("PairClass") == "стена+фасонина" and "отвод воздуховода" in text


def is_transition_fitting(row: dict[str, Any]) -> bool:
    return row.get("PairClass") == "стена+фасонина" and "переход" in path_text(row)


def is_contributing_member(row: dict[str, Any]) -> bool:
    if is_round_damper(row):
        return False
    return row.get("PairClass") in {
        "воздуховод+стена",
        "стена+фасонина",
        "клапан+стена",
    }


def is_wall_pipe(row: dict[str, Any]) -> bool:
    return row.get("PairClass") in {"стена+труба", "изоляция трубы+стена"}


def pipe_opening_size(row: dict[str, Any]) -> float | None:
    """Diameter to use after geometry proves this is a normal wall passage.

    For pipe insulation the Revit Size value often describes the insulation
    layer (for example, ø25), not the insulated assembly. Its geometric outer
    diameter is the safer opening estimate.
    """
    section = as_float(row.get("SectionMax"))
    outer = as_float(row.get("MepMinEdge"))
    if row.get("PairClass") == "изоляция трубы+стена":
        values = [value for value in (section, outer) if value is not None]
        return max(values) if values else None
    return section if section is not None else outer


def is_same_run_duct_and_inline_valve(rows: list[dict[str, Any]]) -> bool:
    """One straight duct plus one same-size inline valve is one passage.

    Navisworks may report the duct body and its inline valve as two clashes
    against the same wall. Their equal nominal sections must not be summed as
    if they were parallel members of a bundle.
    """
    if len(rows) != 2:
        return False

    ducts = [row for row in rows if row.get("PairClass") == "воздуховод+стена"]
    valves = [row for row in rows if row.get("PairClass") == "клапан+стена"]
    if len(ducts) != 1 or len(valves) != 1:
        return False

    duct_min = as_float(ducts[0].get("SectionMin"))
    duct_max = as_float(ducts[0].get("SectionMax"))
    valve_min = as_float(valves[0].get("SectionMin"))
    valve_max = as_float(valves[0].get("SectionMax"))
    return (
        duct_min is not None
        and duct_max is not None
        and duct_min == valve_min
        and duct_max == valve_max
    )


def bound_center_mm(bound: dict[str, Any] | None) -> list[float] | None:
    if not bound:
        return None
    return [
        (bound["Min"][axis] + bound["Max"][axis]) * 0.5 * 304.8
        for axis in ("X", "Y", "Z")
    ]


def spread(points: list[list[float]]) -> list[float]:
    if not points:
        return []
    return [
        round(max(point[i] for point in points) - min(point[i] for point in points), 1)
        for i in range(3)
    ]


def as_float(value: Any) -> float | None:
    try:
        if value is None or isinstance(value, bool):
            return None
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def base_decision(row: dict[str, Any], threshold: float) -> tuple[str, str]:
    pair = str(row.get("PairClass") or "")
    section = as_float(row.get("SectionMax"))

    if is_wall_pipe(row):
        opening_size = pipe_opening_size(row)
        angle = as_float(row.get("WallNormalAngle"))
        end_distance = as_float(row.get("WallEndDistance"))
        end_face_distance = as_float(row.get("WallEndFaceDistance"))
        wall_thickness = as_float(row.get("WallPlanThickness"))

        if not row.get("GeomOk") or angle is None or end_distance is None:
            return "Uncertain", "wall_pipe_geometry_unreliable"

        face_margin = (opening_size or 0.0) * 0.5
        if end_face_distance is not None and end_face_distance <= face_margin:
            return "Active", "pipe_intersects_wall_end_or_reveal_face"

        end_margin = max(wall_thickness or 0.0, (opening_size or 0.0) * 0.5)
        if end_distance <= end_margin:
            return "Active", "pipe_intersects_wall_end"

        # At 45 degrees the axial component along the wall plane becomes
        # greater than the component through its normal. This is no longer a
        # normal opening through the broad face of the wall.
        if angle > 45.0:
            return "Active", "pipe_runs_longitudinally_in_wall"

        if opening_size is not None and opening_size > threshold:
            return "Active", "single_section_over_threshold"
        return "Approved", "normal_wall_passage_in_tolerance"

    if "отделка" in pair:
        if section is not None and section > threshold and row.get("Through") is True:
            return "Active", "oversize_duct_through_finish_layer"
        return "Approved", "finish_layer_no_separate_opening"

    if is_round_damper(row):
        struct = as_float(row.get("StructThick"))
        dist = as_float(row.get("DistMm"))
        if (
            section is not None
            and section > threshold
            and struct
            and dist is not None
            and dist <= struct * ROUND_DAMPER_SHALLOW_DEPTH_RATIO
        ):
            return "Active", "oversize_round_damper_shallow_wall_cut"
        return "Approved", "round_damper_bbox_not_nominal_opening"

    if is_rectangular_damper(row) and section is not None and section > threshold:
        return "Active", "rectangular_damper_over_threshold"

    if is_round_elbow(row):
        return "Active", "duct_elbow_body_in_wall"

    if is_transition_fitting(row):
        return "Active", "transition_fitting_body_in_wall"

    if section is not None and section > threshold:
        return "Active", "single_section_over_threshold"

    return "Approved", "single_section_in_tolerance"


def required_evidence(row: dict[str, Any]) -> list[str]:
    pair = row.get("PairClass")
    supported = {"стена+труба", "изоляция трубы+стена", "воздуховод+стена",
                 "стена+фасонина", "клапан+стена", "воздуховод+отделка",
                 "изоляция воздуховода+стена"}
    if pair not in supported:
        return ["supported_pair_rule"]
    missing = []
    size = pipe_opening_size(row) if is_wall_pipe(row) else as_float(row.get("SectionMax"))
    if size is None or size <= 0:
        missing.append("opening_size")
    if pair == "изоляция воздуховода+стена":
        if row.get("SectionSource") != "Object/Duct Size":
            missing.append("Object/Duct Size")
        thickness = as_float(row.get("InsulationThicknessMm"))
        if thickness is None or thickness <= 0:
            missing.append("InsulationThicknessMm")
    if is_wall_pipe(row):
        if row.get("GeomOk") is not True:
            missing.append("reliable_wall_geometry")
        for key in ("WallNormalAngle", "WallEndDistance", "WallEndFaceDistance",
                    "WallPlanThickness"):
            value = as_float(row.get(key))
            if value is None or value < 0:
                missing.append(key)
    elif pair == "воздуховод+отделка" and row.get("Through") not in (True, False):
        missing.append("Through")
    elif is_round_damper(row):
        for key in ("StructThick", "DistMm"):
            if as_float(row.get(key)) is None:
                missing.append(key)
    elif pair == "клапан+стена" and not is_rectangular_damper(row):
        missing.append("accessory_kind")
    elif pair == "стена+фасонина" and not (is_round_elbow(row) or is_transition_fitting(row)):
        missing.append("fitting_kind")
    return missing


def valid_bound(bound: Any) -> bool:
    if not isinstance(bound, dict):
        return False
    for axis in ("X", "Y", "Z"):
        low = as_float((bound.get("Min") or {}).get(axis))
        high = as_float((bound.get("Max") or {}).get(axis))
        if low is None or high is None or low > high:
            return False
    return True


def classify(rows: list[dict[str, Any]], threshold: float | None = None,
             mode: str = "wall-openings", discipline: str | None = None) -> dict[str, tuple[str, str]]:
    if mode not in {"wall-openings", "ceiling-services", "slab-openings"}:
        raise ValueError("Unknown classifier mode")
    if mode in {"wall-openings", "slab-openings"} and (as_float(threshold) is None or threshold <= 0):
        raise ValueError("Opening threshold must be positive and finite")
    names = [row.get("Name") for row in rows]
    if any(not isinstance(name, str) or not name.strip() for name in names):
        raise ValueError("Every result must have a nonempty Name")
    if len(set(names)) != len(names):
        raise ValueError("Duplicate result names cannot be classified unambiguously")
    if mode == "slab-openings":
        return classify_slabs(rows, threshold)
    if mode == "ceiling-services":
        if discipline not in DISCIPLINES:
            raise ValueError("Ceiling services require OV, VK, PT, EOM or SS discipline")
        return {row["Name"]: classify_ceiling(row, discipline) for row in rows}
    evidence = {row["Name"]: required_evidence(row) for row in rows}
    groups = defaultdict(list)
    for row in rows:
        if group_key(row.get("Group")):
            groups[group_key(row["Group"])].append(row)
    for members in groups.values():
        insulation = [row for row in members if row.get("PairClass") == "изоляция воздуховода+стена"]
        ducts = [row for row in members if row.get("PairClass") == "воздуховод+стена"]
        # A shell is not a second service. Require an unambiguous host before
        # evaluating several insulation members as a common opening.
        if (len(insulation) > 1 or len(ducts) > 1) and any(duct_insulation_host(row, ducts) is None for row in insulation):
            for row in members:
                evidence[row["Name"]].append("insulation_host_for_bundle")
        contributors = [row for row in members if is_contributing_member(row)]
        if len(contributors) < 2:
            continue
        incomplete = any(evidence[row["Name"]] for row in members)
        if not is_same_run_duct_and_inline_valve(contributors):
            incomplete |= any(not valid_bound(row.get("Bound")) for row in members)
        if incomplete:
            for row in members:
                evidence[row["Name"]].append("complete_group_evidence")
    eligible = [row for row in rows if not evidence[row["Name"]]]
    decisions = _classify_supported(eligible, threshold)
    for name, missing in evidence.items():
        if missing:
            decisions[name] = ("Uncertain", "missing:" + ",".join(missing))
    return decisions


def duct_insulation_host(insulation: dict[str, Any], ducts: list[dict[str, Any]]) -> dict[str, Any] | None:
    def service_bound(row):
        side = 1 if row.get("Class2") == "стена" else 2
        return row.get(f"Item{side}Bound")

    outer = service_bound(insulation)
    thickness = as_float(insulation.get("InsulationThicknessMm"))
    if not valid_bound(outer) or thickness is None:
        return None
    sizes = [as_float(insulation.get(key)) for key in ("SectionMin", "SectionMax")]
    if any(value is None for value in sizes):
        return None
    matches = []
    for duct in ducts:
        inner = service_bound(duct)
        dimensions = [as_float(duct.get(key)) for key in ("SectionMin", "SectionMax")]
        if not valid_bound(inner) or any(value is None for value in dimensions):
            continue
        if any(abs(a - 2 * thickness - b) > 2 for a, b in zip(sizes, dimensions)):
            continue
        if all(outer["Min"][axis] - 2 / 304.8 <= inner["Min"][axis]
               and outer["Max"][axis] + 2 / 304.8 >= inner["Max"][axis]
               for axis in ("X", "Y", "Z")):
            matches.append(duct)
    return matches[0] if len(matches) == 1 else None


def _classify_supported(rows: list[dict[str, Any]], threshold: float) -> dict[str, tuple[str, str]]:
    by_group: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        key = group_key(row.get("Group"))
        if key:
            by_group[key].append(row)

    group_active: dict[str, str] = {}
    near_coincident_mixed_groups: set[str] = set()
    same_run_inline_groups: set[str] = set()
    for key, members in by_group.items():
        contributors = [row for row in members if is_contributing_member(row)]
        if len(contributors) < 2:
            continue

        if is_same_run_duct_and_inline_valve(contributors):
            same_run_inline_groups.add(key)
            continue

        outer_sections = {}
        ducts = [row for row in contributors if row.get("PairClass") == "воздуховод+стена"]
        for member in members:
            if member.get("PairClass") != "изоляция воздуховода+стена":
                continue
            host = duct_insulation_host(member, ducts)
            if host is not None:
                name = host["Name"]
                outer_sections[name] = max(outer_sections.get(name, 0), as_float(member.get("SectionMax")) or 0)
        sections = [
            max(as_float(row.get("SectionMax")) or as_float(row.get("MepMinEdge")) or 0.0,
                outer_sections.get(row["Name"], 0))
            for row in contributors
        ]
        sections = [section for section in sections if section > 0]
        if len(sections) < 2:
            continue

        centers = [bound_center_mm(row.get("Bound")) for row in contributors]
        centers = [center for center in centers if center]
        spreads = sorted(spread(centers))
        if len(spreads) != 3:
            continue

        section_sum = sum(sections)
        max_section = max(sections)
        min_section = min(sections)
        mid_spread = spreads[1]

        if (
            len(contributors) == 2
            and max_section >= threshold + 50.0
            and min_section < max_section
            and mid_spread <= max_section * 0.1
        ):
            near_coincident_mixed_groups.add(key)

        effective_opening = max(section_sum, max_section + mid_spread)
        direct_over = any(section > threshold for section in sections)
        separated_enough = mid_spread > max_section * 0.5
        not_far_apart = mid_spread <= section_sum * 3.0
        enough_members_for_mixed_group = (not direct_over) or len(contributors) >= 3
        if (
            separated_enough
            and not_far_apart
            and enough_members_for_mixed_group
            and effective_opening > threshold + ROUNDING_NOISE_MM
        ):
            group_active[key] = "bundle_effective_opening_over_threshold"
            continue

        round_dampers = [row for row in members if is_round_damper(row)]
        if round_dampers and section_sum <= threshold + ROUNDING_NOISE_MM:
            all_centers = [bound_center_mm(row.get("Bound")) for row in contributors + round_dampers]
            all_centers = [center for center in all_centers if center]
            all_spreads = sorted(spread(all_centers))
            if len(all_spreads) == 3:
                all_mid_spread = all_spreads[1]
                if threshold + ROUNDING_NOISE_MM < all_mid_spread <= section_sum * 3.0:
                    group_active[key] = "round_damper_center_confirms_common_opening"

    decisions: dict[str, tuple[str, str]] = {}
    for index, row in enumerate(rows):
        name = str(row.get("Name") or f"row-{index}")
        key = group_key(row.get("Group"))
        if key and key in same_run_inline_groups:
            decision, reason = base_decision(row, threshold)
            decisions[name] = (
                (decision, "same_run_duct_inline_valve_in_tolerance")
                if decision == "Approved"
                else (decision, reason)
            )
        elif key and key in near_coincident_mixed_groups and row.get("PairClass") == "воздуховод+стена":
            decisions[name] = ("Approved", "near_coincident_mixed_group_without_common_opening_evidence")
        elif key and key in group_active:
            decisions[name] = ("Active", group_active[key])
        else:
            decisions[name] = base_decision(row, threshold)
    return decisions


def load_rows(path: str | None) -> list[dict[str, Any]]:
    if not path or path == "-":
        text = sys.stdin.read()
    else:
        with open(path, encoding="utf-8-sig") as source:
            text = source.read()
    payload = json.loads(text)
    if isinstance(payload, dict) and "result" in payload:
        payload = payload["result"]
    if not isinstance(payload, list):
        raise SystemExit("Expected a JSON array or an object with a 'result' array.")
    if any(not isinstance(row, dict) for row in payload):
        raise SystemExit("Every result must be a JSON object.")
    return payload


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8-sig")
    parser = argparse.ArgumentParser()
    parser.add_argument("json_path", nargs="?", default="-")
    parser.add_argument("--mode", choices=("wall-openings", "ceiling-services", "slab-openings"), default="wall-openings")
    parser.add_argument("--mep-discipline", choices=sorted(DISCIPLINES))
    parser.add_argument("--opening-min-edge-mm", type=float)
    parser.add_argument("--compare-status", action="store_true")
    parser.add_argument("--include-names", action="store_true")
    args = parser.parse_args()
    if args.mode in {"wall-openings", "slab-openings"} and (as_float(args.opening_min_edge_mm) is None or args.opening_min_edge_mm <= 0):
        parser.error("opening modes require a positive --opening-min-edge-mm")
    if args.mode == "ceiling-services":
        if not args.mep_discipline:
            parser.error("ceiling-services requires --mep-discipline")
        if args.opening_min_edge_mm is not None:
            parser.error("ceiling-services does not use an opening threshold")

    rows = load_rows(args.json_path)
    decisions = classify(rows, args.opening_min_edge_mm, args.mode, args.mep_discipline)

    decision_counts = Counter(f"{decision}::{reason}" for decision, reason in decisions.values())
    output: dict[str, Any] = {
        "provenance": provenance(rows, args.opening_min_edge_mm, mode=args.mode, discipline=args.mep_discipline),
        "total": len(rows),
        "opening_min_edge_mm": args.opening_min_edge_mm,
        "decision_counts": dict(decision_counts),
        "results": [
            {"name": row["Name"], "group": group_key(row.get("Group")),
             "current_status": row.get("Status"),
             "decision": decisions[row["Name"]][0],
             "reason": decisions[row["Name"]][1],
             "missing_metrics": (decisions[row["Name"]][1][8:].split(",")
                                 if decisions[row["Name"]][1].startswith("missing:") else []),
             "basis": "classifier"}
            for row in rows
        ],
    }

    if args.compare_status:
        mismatches = []
        expected_counts = Counter()
        for index, row in enumerate(rows):
            name = str(row.get("Name") or f"row-{index}")
            expected = status_bucket(row.get("Status"))
            predicted, reason = decisions[name]
            expected_counts[expected] += 1
            if predicted not in {"Uncertain", "Excluded"} and expected != predicted:
                item = {
                    "expected": expected,
                    "predicted": predicted,
                    "reason": reason,
                    "pair": row.get("PairClass"),
                    "section": row.get("SectionMax"),
                    "mep": row.get("MepMinEdge"),
                    "grouped": bool(group_key(row.get("Group"))),
                }
                if args.include_names:
                    item["name"] = name
                mismatches.append(item)
        output["expected_status_counts"] = dict(expected_counts)
        output["mismatch_count"] = len(mismatches)
        output["mismatches"] = mismatches

    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
