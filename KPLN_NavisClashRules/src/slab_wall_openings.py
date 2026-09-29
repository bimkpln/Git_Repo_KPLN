"""Wall-style review for slab/pipe pairs, using independently measured slab faces.

No cylinder fitting or mesh-topology requirement. Axis estimation has the same
scope as the wall workflow: sufficiently elongated bare pipes, with raw item
bounds. Inclined slabs require a signed pipe axis and are left uncertain here.
"""
import math
from collections import defaultdict

FT2MM = 304.8


def number(value):
    if isinstance(value, bool):
        return None
    try:
        value = float(value)
        return value if math.isfinite(value) else None
    except (TypeError, ValueError):
        return None


def box(value):
    if not isinstance(value, dict):
        return None
    try:
        low = [number(value['Min'][k]) for k in 'XYZ']
        high = [number(value['Max'][k]) for k in 'XYZ']
        if any(a is None or b is None or b <= a for a, b in zip(low, high)):
            return None
        return low, high
    except (KeyError, TypeError):
        return None


def evidence(row):
    out = {'source': 'wall-style-item-axis+slab-local-faces-1'}
    if row.get('PairClass') != 'перекрытие+труба':
        return out, 'supported_slab_pipe_pair'
    side = 1 if row.get('Class1') == 'труба' else 2
    pipe_box = box(row.get(f'Item{side}Bound'))
    slab_box = box(row.get(f'Item{3-side}Bound'))
    if pipe_box is None or slab_box is None:
        return out, 'raw_item_bounds'
    if row.get(f'Item{side}BoundSource') != 'leaf':
        return out, 'single_pipe_item_bounds'
    section = number(row.get('SectionMax'))
    if section is None or section <= 0:
        return out, 'nominal_pipe_section'
    dims = [(b-a)*FT2MM for a, b in zip(*pipe_box)]
    # Existing wall _axis algorithm, kept explicit as an estimate (not mesh).
    narrow = min(dims)
    components = [max(d-narrow, 0) for d in dims]
    length = math.sqrt(sum(d*d for d in components))
    out.update(pipe_dimensions_mm=dims, nominal_section_mm=section,
               axis_source='wall-style-item-bounds', pipe_box=pipe_box)
    if narrow + 1e-3 < section or length <= 0 or length/narrow < 2:
        return out, 'reliable_direction_of_short_pipe'
    axis = [d/length for d in components]
    out['pipe_axis_estimate'] = axis
    slab = row.get('PipeSlabGeometry') or {}
    local = slab.get('LocalFaces') or {}
    out.update(pipe_id=row.get('PipeId') or slab.get('PipeId'),
               slab_id=row.get('SlabId') or slab.get('SlabId'))
    if local.get('Source') != 'slab-local-faces-1' or local.get('Usable') is not True:
        return out, 'independent_local_slab_faces'
    normal = local.get('Normal') or {}
    normal = [number(normal.get(k)) for k in 'XYZ']
    if any(x is None for x in normal) or abs(sum(x*x for x in normal)-1) > 1e-5:
        return out, 'slab_normal'
    if abs(normal[2]) < 1-1e-6:
        return out, 'signed_pipe_axis_for_inclined_slab'
    broad = number(local.get('NearestBroadFace'))
    edge = number(local.get('NearestEdgeFace'))
    thickness = number(local.get('Thickness'))
    if broad is None or edge is None or thickness is None or min(broad, edge) < 0 or thickness <= 0:
        return out, 'local_slab_face_distances'
    if not out['pipe_id'] or not out['slab_id']:
        return out, 'pipe_and_slab_identity'
    out.update(slab_normal=normal,
               normal_angle_deg=math.degrees(math.acos(min(1, abs(axis[2])))),
               nearest_broad_face_mm=broad*FT2MM,
               nearest_edge_face_mm=edge*FT2MM,
               slab_thickness_mm=thickness*FT2MM)
    return out, None


def classify_slabs_wall_like(rows, threshold):
    decisions, data, groups = {}, {}, defaultdict(list)
    for row in rows:
        name = row['Name']
        if str(row.get('Status', '')).upper().endswith(('RESOLVED', 'REVIEWED')):
            decisions[name] = ('Excluded', 'preserve_workflow_status')
            continue
        info, missing = evidence(row)
        data[name] = info
        if row.get('GroupId'):
            groups[row['GroupId']].append(row)
        if missing:
            decisions[name] = ('Uncertain', 'missing:' + missing)
        elif info['nearest_edge_face_mm'] <= info['nominal_section_mm']/2:
            decisions[name] = ('Active', 'pipe_intersects_slab_edge_or_reveal_face')
        elif info['normal_angle_deg'] > 45:
            decisions[name] = ('Active', 'pipe_runs_longitudinally_in_slab')
        elif info['nearest_broad_face_mm'] > info['nominal_section_mm']/2:
            decisions[name] = ('Uncertain', 'missing:normal_broad_face_passage')
        elif info['nominal_section_mm'] > threshold:
            decisions[name] = ('Active', 'single_section_over_threshold')
        elif row.get('GroupIdentityKnown') is not True:
            decisions[name] = ('Uncertain', 'missing:stable_group_identity')
        else:
            decisions[name] = ('Approved', 'normal_slab_passage_in_tolerance')

    for members in groups.values():
        if any(decisions[r['Name']][0] == 'Uncertain' for r in members):
            for r in members:
                if decisions[r['Name']][0] == 'Approved':
                    decisions[r['Name']] = ('Uncertain', 'missing:group_member_geometry')
            continue
        # Only transverse members can contribute to a common slab opening.
        # A longitudinal route has its own Active decision, never a duplicate
        # of a small vertical penetration in the same user-created group.
        contributors = [r for r in members if data[r['Name']].get('normal_angle_deg', 90) <= 45]
        for i, a in enumerate(contributors):
            ai = data[a['Name']]
            for b in contributors[:i]:
                bi = data[b['Name']]
                if ai['slab_id'] != bi['slab_id'] or ai['pipe_id'] == bi['pipe_id']:
                    continue
                # For vertical pipes this supplies a conservative separation
                # proof. AABB surface gap is a LOWER bound on actual gap; its
                # transverse width is an UPPER bound on the circular diameter.
                # Thus gap >= upper diameter proves the user's strict bundle
                # condition false, even for boxes inflated by model rotation.
                vertical = max(ai['normal_angle_deg'], bi['normal_angle_deg']) < .1
                ab, bb = ai['pipe_box'], bi['pipe_box']
                lower_gap = math.hypot(*[max(0, ab[0][k]-bb[1][k], bb[0][k]-ab[1][k])*FT2MM for k in (0, 1)])
                upper_diameter = max(ai['pipe_dimensions_mm'][:2] + bi['pipe_dimensions_mm'][:2])
                if vertical and lower_gap + 1e-6 >= upper_diameter:
                    continue
                # Never invent a bundle from proximity alone. Existing measured
                # section mode can resolve these remaining candidates.
                for r in (a, b):
                    if decisions[r['Name']][0] == 'Approved':
                        decisions[r['Name']] = ('Uncertain', 'missing:local_bundle_sections')
    return decisions
