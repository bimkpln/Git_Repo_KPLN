import unittest
from fractions import Fraction as F
from opening_side_intervals import measure,state,outward_floats

def points(xs):return [(x,x) for x in xs]

class SideIntervalTests(unittest.TestCase):
    def test_translation_does_not_cancel(self):
        r=measure(points([0,100,0,100]),points([10,110,0,100]),15,15)
        self.assertEqual(r['technical_sum4'],(20,20));self.assertEqual(r['technical_state'],'over')
    def test_signed_expansion_is_per_side(self):
        r=measure(points([0,100,0,100]),points([-200,300,-200,300]),15,15)
        self.assertEqual(r['KR_outward'],[(200,200)]*4)
    def test_reduction_is_negative(self):
        r=measure(points([0,100,0,100]),points([1,99,1,99]),15,15)
        self.assertEqual(r['KR_outward'],[(-1,-1)]*4)
    def test_mixed_expansion_and_reduction(self):
        r=measure(points([0,100,0,100]),points([-100,99,0,100]),15,15)
        self.assertEqual(r['KR_outward'][:2],[(100,100),(-1,-1)])
    def test_equality_is_allowed(self):self.assertEqual(state((15,15),15),'within')
    def test_no_epsilon_beyond_limit(self):self.assertEqual(state((15+F(1,2**100),)*2,15),'over')
    def test_straddling_limit_is_unknown(self):self.assertEqual(state((14,16),15),'uncertain')
    def test_pair_rule_differs_from_four_sides(self):
        r=measure(points([0,100,0,100]),points([-5,105,-5,105]),15,15)
        self.assertEqual((r['technical_state'],r['architectural_state']),('over','within'))
    def test_uncertain_zero_direction_retained(self):
        r=measure(points([0,100,0,100]),[(-1,1),(100,100),(0,0),(100,100)],15,15)
        self.assertEqual(r['KR_outward'][0],(-1,1));self.assertEqual(r['absolute_edges'][0],(0,1))
    def test_serialization_cannot_shrink_bounds(self):
        a=15-F(1,2**100);b=15+F(1,2**100);lo,hi=outward_floats((a,b))
        self.assertLessEqual(F(lo),a);self.assertGreaterEqual(F(hi),b)
    def test_reversed_interval_rejected(self):
        with self.assertRaises(ValueError):state((1,0),15)
    def test_wrong_side_count_rejected(self):
        with self.assertRaises(ValueError):measure(points([0,1]),points([0,1]),15,15)

if __name__=='__main__':unittest.main()
