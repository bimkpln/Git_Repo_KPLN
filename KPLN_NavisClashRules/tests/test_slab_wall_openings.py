import math
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from classify_opening_clashes import classify


def bounds(low, high):
    return {name: dict(zip('XYZ', [x/304.8 for x in values]))
            for name, values in (('Min', low), ('Max', high))}


def row(name='a', diameter=100, x=0, group=None):
    return dict(Name=name, Status='Active', PairClass='перекрытие+труба',
                Class1='перекрытие', Class2='труба', SectionMax=diameter,
                Item1Bound=bounds([-1500,-1500,0], [1500,1500,200]),
                Item2Bound=bounds([x-diameter/2,-diameter/2,-100], [x+diameter/2,diameter/2,600]),
                Item2BoundSource='leaf', GroupId=group, GroupIdentityKnown=True,
                PipeSlabGeometry=dict(Usable=False, PipeId=name, SlabId='s',
                    LocalFaces=dict(Source='slab-local-faces-1', Usable=True,
                        Normal=dict(X=0,Y=0,Z=1), Thickness=200/304.8,
                        NearestBroadFace=0, NearestEdgeFace=1000/304.8)))


def decide(rows, threshold=150):
    return {n:d for n,(d,r) in classify(rows, threshold, 'slab-wall-openings').items()}


class SlabWallReviewTests(unittest.TestCase):
    def test_closed_cylinder_not_needed_for_normal_passage(self):
        self.assertEqual(decide([row()]), {'a':'Approved'})

    def test_size_boundary_matches_walls(self):
        for d, expected in ((149,'Approved'),(150,'Approved'),(151,'Active')):
            self.assertEqual(decide([row(diameter=d)])['a'], expected)

    def test_edge_or_hole_reveal_overrides_small_size(self):
        r=row(diameter=25)
        r['PipeSlabGeometry']['LocalFaces']['NearestEdgeFace']=10/304.8
        self.assertEqual(decide([r])['a'], 'Active')

    def test_horizontal_pipe_is_longitudinal(self):
        r=row(diameter=25)
        r['Item2Bound']=bounds([-300,-100,60],[300,100,85])
        self.assertEqual(decide([r])['a'], 'Active')

    def test_small_crossing_with_longitudinal_neighbor_is_separate(self):
        a=row(group='g',diameter=25); b=row('b',group='g',diameter=25)
        b['Item2Bound']=bounds([-300,-100,60],[300,100,85])
        self.assertEqual(decide([a,b]),{'a':'Approved','b':'Active'})

    def test_missing_reveal_distance_not_approved(self):
        r=row();r['PipeSlabGeometry']['LocalFaces']['NearestEdgeFace']=None
        self.assertEqual(decide([r])['a'],'Uncertain')

    def test_short_pipe_does_not_get_longest_box_axis(self):
        r=row();r['Item2Bound']=bounds([-50,-50,-10],[50,50,10])
        self.assertEqual(decide([r])['a'],'Uncertain')

    def test_inclined_slab_requires_signed_axis(self):
        r=row();r['PipeSlabGeometry']['LocalFaces']['Normal']=dict(X=.6,Y=0,Z=.8)
        self.assertEqual(decide([r])['a'],'Uncertain')

    def test_group_distance_and_exact_gap_boundary(self):
        self.assertEqual(decide([row(group='g'),row('b',x=200,group='g')]),{'a':'Approved','b':'Approved'})
        self.assertEqual(decide([row(group='g'),row('b',x=199,group='g')]),{'a':'Uncertain','b':'Uncertain'})

    def test_duplicate_pipe_on_different_slabs_is_not_bundle(self):
        a=row(group='g');b=row('b',group='g')
        b['PipeSlabGeometry'].update(PipeId='a',SlabId='other')
        self.assertEqual(decide([a,b]),{'a':'Approved','b':'Approved'})

    def test_workflow_exclusion_before_group(self):
        a=row(group='g');b=row('b',group='g');b['Status']='Reviewed'
        self.assertEqual(decide([a,b]),{'a':'Approved','b':'Excluded'})

    def test_threshold_required(self):
        with self.assertRaises(ValueError):
            classify([row()], 0, 'slab-wall-openings')


if __name__ == '__main__':
    unittest.main()
