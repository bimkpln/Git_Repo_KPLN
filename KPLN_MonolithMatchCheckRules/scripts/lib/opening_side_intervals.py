"""Exact interval arithmetic after geometric side correspondence is proved.

Input order is left/right/bottom/top. This module neither pairs openings nor
proves contour containment, throughness, lintel roles or an assembly's voids.
"""
from fractions import Fraction
import math


def interval(value):
    if len(value)!=2:raise ValueError('Two interval endpoints required')
    a,b=(Fraction(x) for x in value)
    if a>b:raise ValueError('Reversed interval')
    return a,b


def subtract(a,b):
    a,b=interval(a),interval(b)
    return a[0]-b[1],a[1]-b[0]


def absolute(a):
    lo,hi=interval(a)
    return (Fraction(0) if lo<=0<=hi else min(abs(lo),abs(hi)),max(abs(lo),abs(hi)))


def state(value,limit):
    lo,hi=interval(value);limit=Fraction(limit)
    if limit<0:raise ValueError('A nonnegative tolerance is required')
    return 'over' if lo>limit else 'within' if hi<=limit else 'uncertain'


def measure(ar,kr,technical_limit,architectural_limit):
    if len(ar)!=4 or len(kr)!=4:raise ValueError('Four corresponding sides required')
    ar=list(map(interval,ar));kr=list(map(interval,kr))
    signed=[subtract(ar[i],kr[i]) if i%2==0 else subtract(kr[i],ar[i]) for i in range(4)]
    edges=list(map(absolute,signed))
    total=tuple(sum(x[j] for x in edges) for j in range(2))
    pairs=[tuple(edges[i][j]+edges[i+1][j] for j in range(2)) for i in [0,2]]
    worst=tuple(max(x[j] for x in pairs) for j in range(2))
    return {'KR_outward':signed,'absolute_edges':edges,'technical_sum4':total,
            'opposite_pairs':pairs,'max_opposite_pair':worst,
            'technical_state':state(total,technical_limit),
            'architectural_state':state(worst,architectural_limit)}


def outward_floats(value):
    """JSON endpoints rounded outwards, retaining exact values separately."""
    a,b=interval(value);lo=float(a);hi=float(b)
    if not math.isfinite(lo) or not math.isfinite(hi):raise ValueError('Non-finite serialized interval')
    if Fraction(lo)>a:lo=math.nextafter(lo,-math.inf)
    if Fraction(hi)<b:hi=math.nextafter(hi,math.inf)
    return [lo,hi]
