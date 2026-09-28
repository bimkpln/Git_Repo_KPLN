"""Pipe/slab policy. All geometry is measured by the bridge, never by AABB axes."""
import math
from collections import defaultdict


def number(value):
    if isinstance(value, bool):
        return None
    try:
        value = float(value)
        return value if math.isfinite(value) else None
    except (ValueError, TypeError):
        return None


def vector(value):
    if not isinstance(value, (list, tuple)) or len(value) != 3:
        return None
    result = [number(x) for x in value]
    return result if all(x is not None for x in result) else None


def dot(a, b):
    return sum(x*y for x, y in zip(a, b))


def sub(a, b):
    return [x-y for x, y in zip(a, b)]


def cross(a, b):
    return [a[1]*b[2]-a[2]*b[1], a[2]*b[0]-a[0]*b[2], a[0]*b[1]-a[1]*b[0]]


def basis(normal):
    u = cross(normal, [0, 0, 1] if abs(normal[2]) < .9 else [1, 0, 0])
    length = math.sqrt(dot(u, u))
    u = [x/length for x in u]
    return u, cross(normal, u)


def cross2(a, b, c):
    return (b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0])


def hull(points):
    points = sorted(set(tuple(p) for p in points))
    lower, upper = [], []
    for seq, output in ((points, lower), (reversed(points), upper)):
        for p in seq:
            while len(output) >= 2 and cross2(output[-2], output[-1], p) <= 0:
                output.pop()
            output.append(p)
    return lower[:-1] + upper[:-1]


def project(points, origin, u, v):
    return hull([(dot(sub(p, origin), u), dot(sub(p, origin), v)) for p in points])


def polygon_gap(a, b):
    # Separating axis test first: intersecting/crossing polygons have gap zero,
    # even when neither contains a vertex of the other.
    separated = False
    for poly in (a, b):
        for p, q in zip(poly, poly[1:]+poly[:1]):
            axis = [q[1]-p[1], p[0]-q[0]]
            pa, pb = [dot(x, axis) for x in a], [dot(x, axis) for x in b]
            if max(pa) < min(pb) or max(pb) < min(pa):
                separated = True
    if not separated:
        return 0.0
    best = math.inf
    for first, second in ((a, b), (b, a)):
        for p in first:
            for q, r in zip(second, second[1:]+second[:1]):
                d = sub(r, q)
                length2 = dot(d, d)
                t = min(1, max(0, dot(sub(p, q), d)/length2)) if length2 else 0
                best = min(best, math.hypot(p[0]-q[0]-t*d[0], p[1]-q[1]-t*d[1]))
    return best


def opening_edge(polygons):
    points = hull([p for polygon in polygons for p in polygon])
    best = (math.inf, math.inf)
    for a, b in zip(points, points[1:]+points[:1]):
        dx, dy = b[0]-a[0], b[1]-a[1]
        length = math.hypot(dx, dy)
        if not length:
            continue
        along = [(p[0]*dx+p[1]*dy)/length for p in points]
        across = [(-p[0]*dy+p[1]*dx)/length for p in points]
        x, y = max(along)-min(along), max(across)-min(across)
        best = min(best, (x*y, max(x, y)))
    return best[1]


def geometry(row):
    normal = vector(row.get("SlabNormal"))
    axis = vector(row.get("SlabPipeAxis"))
    angle = number(row.get("SlabNormalAngle"))
    outer = number(row.get("SlabOuterDiameterMm"))
    size = number(row.get("SectionMax"))
    if (row.get("SlabMetricsVersion") != "pipe-slab-mesh-1" or row.get("GeomOk") is not True
            or normal is None or axis is None or angle is None or not 0 <= angle <= 90
            or outer is None or outer <= 0 or size is None or size <= 0
            or not row.get("PipeId") or not row.get("SlabId")
            or type(row.get("SlabEdgeContact")) is not bool
            or type(row.get("SlabFullSectionCovered")) is not bool):
        return None
    if abs(dot(normal, normal)-1) > 1e-5 or abs(dot(axis, axis)-1) > 1e-5:
        return None
    actual = math.degrees(math.acos(min(1, abs(dot(normal, axis)))))
    if abs(actual-angle) > 1e-5:
        return None
    sections = []
    for section in row.get("SlabSections") or []:
        origin = vector(section.get("origin_mm"))
        points = [vector(p) for p in section.get("contour_mm", [])]
        if origin is None or len(points) < 3 or any(p is None for p in points):
            return None
        if any(abs(dot(sub(p, origin), normal)) > .01 for p in points):
            return None
        u, v = basis(normal)
        if len(project(points, origin, u, v)) < 3:
            return None
        sections.append((origin, points))
    if row.get("SlabFullSectionCovered") and not sections:
        return None
    return normal, angle, outer, size, sections


def classify_slabs(rows, threshold):
    decisions, groups = {}, defaultdict(list)
    for row in rows:
        name = row["Name"]
        status = str(row.get("Status", "")).upper()
        if status.endswith(("RESOLVED", "REVIEWED")):
            decisions[name] = ("Excluded", "preserve_workflow_status")
            continue
        if row.get("GroupId"):
            groups[row["GroupId"]].append(row)
        if row.get("PairClass") != "перекрытие+труба":
            decisions[name] = ("Uncertain", "missing:supported_slab_pipe_pair")
            continue
        g = geometry(row)
        if g is None:
            decisions[name] = ("Uncertain", "missing:verified_pipe_axis_and_local_slab_mesh")
        elif row["SlabEdgeContact"]:
            decisions[name] = ("Active", "pipe_intersects_slab_edge_or_reveal")
        elif g[1] > 45 + 1e-7:
            decisions[name] = ("Active", "pipe_runs_along_slab")
        elif abs(g[1]-45) <= 1e-7:
            decisions[name] = ("Uncertain", "missing:review_of_45_degree_boundary")
        elif not row["SlabFullSectionCovered"]:
            decisions[name] = ("Uncertain", "missing:complete_broad_face_passage")
        elif g[3] > threshold:
            decisions[name] = ("Active", "pipe_section_above_opening_threshold")
        elif row.get("GroupIdentityKnown") is not True:
            decisions[name] = ("Uncertain", "missing:stable_group_identity")
        else:
            decisions[name] = ("Approved", "normal_pipe_passage_within_threshold")
    for members in groups.values():
        # Unknown members can conceal a nearby pipe that enlarges the opening.
        if any(geometry(r) is None or r.get("GroupIdentityKnown") is not True for r in members):
            for r in members:
                if decisions[r["Name"]][0] == "Approved":
                    decisions[r["Name"]] = ("Uncertain", "missing:complete_bundle_geometry")
            continue
        by_slab = defaultdict(list)
        for row in members:
            by_slab[row["SlabId"]].append(row)
        for same_slab in by_slab.values():
            apply_bundles(same_slab, decisions, threshold)
    return decisions


def apply_bundles(rows, decisions, threshold):
    nodes = defaultdict(list)
    for row in rows:
        nodes[row["PipeId"]].append(row)
    if len(nodes) < 2:
        return
    representatives = [members[0] for members in nodes.values()]
    base = geometry(representatives[0])
    n = base[0]
    u, v = basis(n)
    # Only full sections in one common slab can establish spacing. No group
    # inheritance of axes, no centers/Bound substitution for outer contours.
    if any(not geometry(r)[4] or abs(dot(n, geometry(r)[0])) < 1-1e-6
           or not r["SlabFullSectionCovered"] for r in representatives):
        for row in rows:
            if decisions[row["Name"]][0] == "Approved":
                decisions[row["Name"]] = ("Uncertain", "missing:bundle_sections_on_common_slab")
        return
    origin = base[4][0][0]
    sections = [[(dot(sub(o, origin), n), project(p, origin, u, v)) for o, p in geometry(r)[4]] for r in representatives]
    adjacency = {i: set() for i in range(len(representatives))}
    for i in adjacency:
        for j in range(i):
            common = [(a, b) for za, a in sections[i] for zb, b in sections[j] if abs(za-zb) <= .01]
            if not common:
                for row in rows:
                    if decisions[row["Name"]][0] == "Approved":
                        decisions[row["Name"]] = ("Uncertain", "missing:bundle_sections_on_common_plane")
                return
            gap = min(polygon_gap(a, b) for a, b in common)
            limit = max(geometry(representatives[i])[2], geometry(representatives[j])[2])
            if gap < limit:  # User-confirmed strict comparison, surface-to-surface.
                adjacency[i].add(j)
                adjacency[j].add(i)
    remaining = set(adjacency)
    while remaining:
        pending, component = [min(remaining)], set()
        while pending:
            i = pending.pop()
            if i in component:
                continue
            component.add(i)
            pending.extend(adjacency[i]-component)
        remaining -= component
        if len(component) < 2:
            continue
        edge = opening_edge([p for i in component for _, p in sections[i]])
        for i in component:
            for row in nodes[representatives[i]["PipeId"]]:
                # Geometry failures and confirmed edge/longitudinal clashes are
                # never downgraded by a bundle's size.
                if decisions[row["Name"]][0] == "Approved" and edge > threshold:
                    decisions[row["Name"]] = ("Active", "bundle_opening_above_threshold")
