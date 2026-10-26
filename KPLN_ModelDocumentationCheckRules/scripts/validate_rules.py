"""Validate every registered scope and referenced layer, with no Revit access."""
from pathlib import Path
from resolve_rules import read_json,resolve
import sys
root=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else Path(__file__).resolve().parents[1]
reg=read_json(root/'registry.json')
resolve(root)
paths={'rules/common.json'}
for department,item in reg['departments'].items():
    if item['path'] in paths:raise ValueError('Reused layer path')
    paths.add(item['path']);resolve(root,department)
for project,item in reg['projects'].items():
    for department,layer in item['departments'].items():
        if department not in reg['departments']:raise ValueError('Project references unknown department')
        if layer['path'] in paths:raise ValueError('Reused layer path')
        paths.add(layer['path']);resolve(root,department,project)
for path in root.glob('departments/**/rules.json'):
    if path.relative_to(root).as_posix() not in paths:raise ValueError('Unregistered layer '+str(path))
for path in root.glob('projects/**/rules.json'):
    if path.relative_to(root).as_posix() not in paths:raise ValueError('Unregistered layer '+str(path))
print('Valid:',len(paths),'registered layers')
