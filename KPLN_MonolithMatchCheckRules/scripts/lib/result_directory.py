"""Project report folders; atomic allocation preserves previous runs."""
import datetime
from pathlib import Path

MOSCOW=datetime.timezone(datetime.timedelta(hours=3),'Europe/Moscow')

def check_mode(config):
    variant=config.get('check_variant','full')
    exclusions=config.get('additional_exclusions',[])
    if variant not in ['full','lite']:raise ValueError('check_variant must be full or lite')
    if not isinstance(exclusions,list) or any(not isinstance(e,str) or not e.strip() for e in exclusions):
        raise ValueError('additional_exclusions must list explicit exclusions')
    if variant=='lite' and not exclusions:raise ValueError('Lite requires explicit additional exclusions')
    if variant=='full' and exclusions:raise ValueError('Additional exclusions require lite, not full')
    return variant,exclusions

def folder_name(config,now=None):
    variant,_=check_mode(config);area=config.get('result_scope_label','')
    if not isinstance(area,str) or not area.strip() or area!=area.strip() or area.endswith('.') or area in ['.','..'] or any(c in area for c in '<>:"/\\|?*') or any(ord(c)<32 for c in area):
        raise ValueError('Explicit filename-safe result_scope_label required (e.g. К2С14_Этажи01-05)')
    now=now or datetime.datetime.now(datetime.timezone.utc)
    if now.tzinfo is None:raise ValueError('Timezone-aware execution time required')
    date=now.astimezone(MOSCOW).strftime('%Y.%m.%d')
    return date+'_Монолит_'+('Фулл' if variant=='full' else 'Лайт')+'_'+area

def allocate(root,config,now=None):
    root=Path(root).resolve();base=folder_name(config,now);root.mkdir(parents=True,exist_ok=True)
    number=1
    while True:
        target=root/(base if number==1 else base+f'_{number:02d}')
        try:target.mkdir();return target
        except FileExistsError:number+=1
