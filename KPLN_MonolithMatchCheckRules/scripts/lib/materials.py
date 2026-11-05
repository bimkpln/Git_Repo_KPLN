"""Evidence-based fragment selection; unknown is not absence of concrete."""
import re

NON_MONOLITH=r'кирпич|газобетон|пенобетон|гипс|утепл|минерал.*ват|пенополистирол|сталь|металл|алюмин|стекл|дерев|керам|стяжк|плитк|линолеум|паркет|краска|сборн|prefab|precast|steel|insulation|masonry|paint'
CONCRETE=r'железобетон|бетон|(?:^|[_\s(])жб(?:$|[_\s)])|concrete'
CAST=r'монолит|cast.in.place|in.situ'

def classify_fragment(fragment,approved_materials=()):
    """Adapter supplies material for this mesh fragment, not a parent wall list.

    approved_materials is a project-specific verified dictionary, never learned
    merely from the fact that a previous NWD had a 'Монолит' filename.
    """
    if fragment.get('excluded_by_hide_set'):
        return {'decision':'exclude','reason':'recursive_hide_search_set'}
    if fragment.get('is_void_cutter'):
        return {'decision':'exclude','reason':'cutter_not_material; actual_void_checked_in_host'}
    material=str(fragment.get('material','')).lower()
    method=str(fragment.get('construction_method','')).lower()
    category=str(fragment.get('category','')).lower()
    role=str(fragment.get('role','')).lower()
    # Explicit purpose/method and exact material refer to this physical fragment.
    if re.search(r'сборн|precast|prefab',method) or role in ['finish','insulation','nonstructural_floor_finish']:
        return {'decision':'exclude','reason':'confirmed_non_monolith_method_or_role'}
    if re.search(NON_MONOLITH,material):return {'decision':'exclude','reason':'non_monolith_fragment_material'}
    if category in ['окна','двери','windows','doors'] and role!='monolithic_threshold':
        return {'decision':'exclude','reason':'nonstructural_window_door_product'}
    if material and material in {x.lower() for x in approved_materials}:
        return {'decision':'include','reason':'project_verified_fragment_material'}
    if not material and fragment.get('verified_original_closure'):
        return {'decision':'include','reason':'source_face_closes_proven_monolith; exact_input_geometry_certificate'}
    if re.search(CONCRETE,material) and (re.search(CAST,material) or re.search(CAST,method)):
        return {'decision':'include','reason':'concrete_and_cast_in_place_evidence'}
    return {'decision':'unresolved','reason':'insufficient_material_or_construction_method; do_not_treat_as_missing'}

def select_fragments(fragments,approved_materials=()):
    records=[{'fragment_id':f['fragment_id'],**classify_fragment(f,approved_materials)} for f in fragments]
    ids=[r['fragment_id'] for r in records]
    if len(ids)!=len(set(ids)):raise ValueError('Fragment identity must be unique')
    return {'complete':not any(r['decision']=='unresolved' for r in records),'records':records,
            'counts':{d:sum(r['decision']==d for r in records) for d in ['include','exclude','unresolved']}}
