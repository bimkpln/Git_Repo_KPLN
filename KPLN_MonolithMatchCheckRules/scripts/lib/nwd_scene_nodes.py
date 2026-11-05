"""Unpack full scene records explicitly; never iterate a JSON envelope's keys."""
def scene_nodes(payload,expected_count):
    if type(expected_count) is not int or expected_count<0:raise ValueError('Explicit nonnegative integer scene count required')
    if not isinstance(payload,dict) or not isinstance(payload.get('nodes'),list):raise ValueError('Full NWD scene envelope with a nodes array is required')
    nodes=payload['nodes']
    if len(nodes)!=expected_count:raise ValueError('Full scene node count differs from the export manifest')
    for index,node in enumerate(nodes):
        if not isinstance(node,dict) or type(node.get('scene_index')) is not int or node.get('scene_index')!=index:raise ValueError('Scene node identity is not contiguous')
        parent=node.get('parent_index')
        if type(parent) is not int or parent>=index or parent< -1:raise ValueError('Invalid full scene hierarchy')
        if not isinstance(node.get('properties'),list):raise ValueError('Scene properties must be an array')
    return nodes
