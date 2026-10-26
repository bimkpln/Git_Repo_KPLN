"""Resolve explicitly selected KPLN rule layers; does not inspect Revit."""
import argparse
import hashlib
import json
from pathlib import Path

def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))

def safe_path(root, relative):
    path=(root/relative).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError('Layer path escapes rules root: '+relative)
    return path

def load_layer(root, relative, expected_scope):
    path=safe_path(root,relative)
    layer=read_json(path)
    if layer.get('schema_version')!=1 or layer.get('scope')!=expected_scope:
        raise ValueError('Invalid schema/scope: '+relative)
    if not isinstance(layer.get('rules'),list):
        raise ValueError('Rules must be a list: '+relative)
    seen=set()
    for rule in layer['rules']:
        for key in ('id','title','condition','disposition','severity','source'):
            if not isinstance(rule.get(key),str) or not rule[key].strip():
                raise ValueError('Missing text field '+key+' in '+relative)
        if rule['id'] in seen:raise ValueError('Duplicate rule id: '+rule['id'])
        seen.add(rule['id'])
        if type(rule.get('active')) is not bool:raise ValueError('active must be boolean')
        if 'overrides' in rule and type(rule['overrides']) is not bool:raise ValueError('overrides must be boolean')
        if rule['disposition'] not in ('allowed','warning','review','violation'):raise ValueError('Unknown disposition')
        if rule['severity'] not in ('none','warning','medium','critical'):raise ValueError('Unknown severity')
        if rule['disposition']=='allowed' and rule['severity']!='none':raise ValueError('Allowed rule must have severity none')
    return layer,hashlib.sha256(path.read_bytes()).hexdigest()

def resolve(root,department=None,project=None,*,common_only=False):
    if type(common_only) is not bool:raise ValueError('common_only must be boolean')
    root=Path(root).resolve()
    registry=read_json(root/'registry.json')
    if registry.get('schema_version')!=1:raise ValueError('Invalid registry version')
    selected=[('rules/common.json',{'level':'common'})]
    warnings=[]
    departments=registry['departments'];projects=registry['projects']
    if common_only:
        warnings.append('По явному ограничению применён только общий слой; локальные слои не читались.')
    elif not department:
        warnings.append('Отдел не задан: применён только общий слой.')
    elif department not in departments:
        warnings.append('Неизвестный отдел '+department+': локальные допуски не применены.')
    else:
        selected.append((departments[department]['path'],dict(level='department',department=department)))
    if common_only:
        pass
    elif not project:
        warnings.append('Проект не задан: проектно-отделовой слой не применён.')
    elif project not in projects:
        warnings.append('Неизвестный проект '+project+': проектные допуски не применены.')
    elif department in departments:
        mapping=projects[project]['departments']
        if department in mapping:
            selected.append((mapping[department]['path'],dict(level='project_department',project=project,department=department)))
        else:
            warnings.append('Для проекта и отдела отдельный слой не зарегистрирован.')
    else:
        warnings.append('Проектный слой требует известного отдела; он не применён.')
    effective={};history={};layers=[]
    for relative,scope in selected:
        layer,digest=load_layer(root,relative,scope)
        layers.append(dict(path=relative,scope=scope,sha256=digest))
        for record in layer['rules']:
            rule=dict(record);key=rule['id'];exists=key in effective
            if exists and not rule.get('overrides'):
                raise ValueError('Explicit overrides=true required for '+key+' in '+relative)
            if rule.get('overrides') and not exists:
                raise ValueError('Cannot override absent rule '+key+' in '+relative)
            reason=rule.get('reason')
            if exists and (not isinstance(reason,str) or not reason.strip()):
                raise ValueError('Override reason required for '+key)
            if not rule['active'] and (not isinstance(reason,str) or not reason.strip()):
                raise ValueError('Disabling a rule requires reason: '+key)
            history.setdefault(key,[]).append(dict(layer=relative,scope=scope,rule=dict(rule)))
            rule['resolved_layer']=relative
            effective[key]=rule
    return dict(version=(root/'VERSION').read_text(encoding='utf-8').strip(),
                context=dict(department=department,project=project),selection_mode='common_only' if common_only else 'applicable_layers',layers=layers,warnings=warnings,
                rules=list(effective.values()),history=history)

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1])
    parser.add_argument('--department');parser.add_argument('--project');parser.add_argument('--output',type=Path)
    parser.add_argument('--common-only',action='store_true',help='Explicitly apply only common rules while preserving known context')
    args=parser.parse_args()
    try:result=resolve(args.root,args.department,args.project,common_only=args.common_only)
    except (ValueError,KeyError,OSError,TypeError) as e:parser.exit(2,str(e)+'\n')
    text=json.dumps(result,ensure_ascii=False,indent=2)+'\n'
    if args.output:args.output.write_text(text,encoding='utf-8')
    else:print(text,end='')
