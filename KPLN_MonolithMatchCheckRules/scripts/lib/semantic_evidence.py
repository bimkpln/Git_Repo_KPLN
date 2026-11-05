"""Declared-role candidates from instance names and scoped property values.

Candidates never constitute a geometry classification or a BEP exemption.
Workset names and the mere presence of a parameter are deliberately excluded.
"""
import re
PREFIX=r'(?<![0-9A-Za-zА-Яа-яЁё])'
PATTERNS={
    'pit':PREFIX+r'приям\w*',
    'shaft':PREFIX+r'шахт\w*',
    'mounting':PREFIX+r'(?:монтажн|установочн)\w*',
    'expansion_joint':PREFIX+r'(?:деформационн|температурн|осадочн)\w*',
    'door_window':PREFIX+r'(?:двер\w*|окн(?:[аоыуе]\w*|(?=[ _-]|$))|оконн\w*|окон(?=[ _-]|$))',
}
ROLE_FIELDS={'Имя','Тип','Семейство','Описание','Комментарии','Назначение','КП_О_Назначение','КП_О_Обозначение','Орг.ГруппаВедомостьРасходаБетона'}

def declared_role_candidates(node):
    values=[{'scope':'node_name','value':str(node.get('name',''))}]
    for p in node.get('properties',[]):
        if p.get('category')!='Объект' or p.get('name') not in ROLE_FIELDS:continue
        value=re.sub(r'^[A-Za-z0-9_ ]+:','',str(p.get('value','')),count=1)
        values.append({'scope':'instance_property','category':p['category'],'property':p['name'],
            'property_name':p.get('property_name'),'value':value})
    return {role:[v for v in values if re.search(pattern,v['value'],re.I)] for role,pattern in PATTERNS.items()}
