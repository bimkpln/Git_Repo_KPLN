"""Add lifecycle UI without changing geometry, IDs or native plugin columns."""
import json

def render(template,entries,plans,summary,profile):
    local=[g for e in entries for g in e.get('local_instances') or [e]]
    if any(not g.get('first_detected') or not g.get('lifecycle_status') for g in local):raise ValueError('History reconciliation must precede report rendering')
    fields=['number','title','grid_location','image','review_image','children','level_name','entry_type','measurement_mm',
        'tolerance_mm','measurement_rule','comment','category','parent_number','primary_plan','previous_numbers','plan_label',
        'first_detected','lifecycle_status','issue_uid']
    data=[{k:g.get(k) for k in fields} for g in local]
    statuses={'new':'Новое','unworked':'Неотработанное','baseline':'Исходная итерация','reopened':'Возобновлено','not_checked':'Не проверялось','resolved':'Устранено'}
    selector='<label>Статус <select id="lifecycleFilter"><option value="">Все</option>'+''.join(f'<option value="{k}">{v}</option>' for k,v in statuses.items())+'</select></label>'
    template=template.replace('<div id="tools">','<div id="tools">'+selector)
    template=template.replace('const searchIndex=', 'const lifecycleLabels='+json.dumps(statuses,ensure_ascii=False)+';\nconst searchIndex=')
    old=".filter(g=>!q||searchIndex.get(g.number).includes(q))"
    new=".filter(g=>(!document.getElementById('lifecycleFilter').value||g.lifecycle_status===document.getElementById('lifecycleFilter').value)&&(!q||searchIndex.get(g.number).includes(q)))"
    if old not in template:raise ValueError('Coordinator template changed; update lifecycle adapter')
    template=template.replace(old,new)
    # Apply the status predicate to cross-plan search fallback too.
    template=template.replace("groups.filter(g=>searchIndex.get(g.number).includes(q))", "groups.filter(g=>(!document.getElementById('lifecycleFilter').value||g.lifecycle_status===document.getElementById('lifecycleFilter').value)&&searchIndex.get(g.number).includes(q))")
    template=template.replace('<b>${esc(g.plan_label||g.number)}</b>', '<b>${esc(g.plan_label||g.number)}</b><span style="display:block">${esc(lifecycleLabels[g.lifecycle_status])} · Впервые в реестре: ${esc(g.first_detected)}</span>')
    template=template.replace('level.onchange=()=>change(level.value);', "document.getElementById('lifecycleFilter').onchange=list;level.onchange=()=>change(level.value);")
    old='ПД · конструкции 15 мм · низ марша 20 мм · проёмы: пары граней 15 мм · технические отверстия: сумма граней 100 мм'
    label='ПД' if profile['stage']=='PD' else 'РД'
    text=f"{label} · конструкции {profile['general_mm']} мм · низ марша {profile['flight_underside_vertical_mm']} мм · проёмы: пары граней {profile['door_window_opposite_pair_mm']} мм · технические отверстия: сумма граней {profile['technical_hole_sum_four_mm']} мм"
    template=template.replace(old,text)
    for token,value in [('__SUMMARY__',summary),('__PLAN_DATA__',plans),('__GROUP_DATA__',data)]:
        template=template.replace(token,json.dumps(value,ensure_ascii=False).replace('</','<\\/'))
    return template
