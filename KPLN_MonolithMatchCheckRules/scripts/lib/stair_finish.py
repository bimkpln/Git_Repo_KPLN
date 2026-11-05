"""Explicit AR finish evidence and continuous local thickness proof."""
import re
import numpy as np
import shapely
from vertical_surface import clip_positive

def explicit_finish(properties):
    text=' '.join(str(properties.get(k,'')) for k in ['Семейство','Тип','Имя','Комментарии']).casefold()
    if re.search(r'\bбез\s+отделк',text):return False
    return bool(re.search(r'\bс\s+отделк|\bотделка\s+(?:плитк|кам|керам)',text))

def finish_region(shape,ar_top,ar_bottom,kr_top,kr_bottom,anchor_tolerance,vertical_error,anchor='bottom'):
    at,ab,kt,kb=[np.asarray(p,float) for p in [ar_top,ar_bottom,kr_top,kr_bottom]]
    if min(anchor_tolerance,vertical_error)<0:raise ValueError('Negative finish proof bound')
    d=(ab-kb) if anchor=='bottom' else (at-kt)
    limit=anchor_tolerance-vertical_error
    if limit<0:return shapely.GeometryCollection()
    aligned=clip_positive(clip_positive(shape,d,-limit),-d,-limit)
    positive=clip_positive(clip_positive(aligned,at-ab,vertical_error+1e-8),kt-kb,vertical_error+1e-8)
    thickness=(at-ab)-(kt-kb)
    return shapely.union_all([clip_positive(positive,thickness,vertical_error+1e-8),clip_positive(positive,-thickness,vertical_error+1e-8)])
