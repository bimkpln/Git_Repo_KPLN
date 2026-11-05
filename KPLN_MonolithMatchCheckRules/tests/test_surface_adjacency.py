import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
from surface_adjacency import partial_edge_links
class EdgeAdjacencyTests(unittest.TestCase):
    def test_T_subdivision_connects(self):
        edges=[[[0,0,0],[10,0,0]],[[0,0,0],[4,0,0]],[[4,0,0],[10,0,0]]]
        links,proof=partial_edge_links(edges,[0,1,2]);self.assertEqual(len(links),2)
        self.assertEqual(sum(r['continuous_overlap_length_mm'] for r in links),10)
    def test_point_contact_does_not_connect(self):
        links,p=partial_edge_links([[[0,0,0],[4,0,0]],[[4,0,0],[10,0,0]]],[0,1]);self.assertEqual(links,[])
    def test_crossing_noncollinear_edges_do_not_connect(self):
        links,p=partial_edge_links([[[0,0,0],[10,0,0]],[[5,-5,0],[5,5,0]]],[0,1]);self.assertEqual(links,[])
    def test_parallel_actual_gap_is_not_connected(self):
        links,p=partial_edge_links([[[0,0,0],[10,0,0]],[[0,0,.001],[10,0,.001]]],[0,1]);self.assertEqual(links,[])
    def test_zero_length_edges_do_not_connect(self):
        links,p=partial_edge_links([[[0,0,0],[10,0,0]],[[5,0,0],[5,0,0]]],[0,1]);self.assertEqual(links,[])
if __name__=='__main__':unittest.main()
