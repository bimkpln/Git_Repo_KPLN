"""Portable source-topology tests, including rejection of actual missing faces."""
import sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from topology_edges import *

V=[(0,0,0),(1,0,0),(0,1,0),(0,0,1)]
F=[(0,2,1),(0,1,3),(1,2,3),(2,0,3)]

class SurfaceTests(unittest.TestCase):
    def test_preserves_oriented_source_triangles(self):
        v,f,c=split_contact_edges(V,F)
        self.assertEqual(triangle_hash(V,F),triangle_hash(v,f))

    def test_touching_bodies_separate_without_new_surface(self):
        vv=V+[(0,0,0),(1,0,0),(0,-1,0),(0,0,-1)]
        ff=F+[tuple(i+4 for i in face) for face in F]
        v=list(dict.fromkeys(vv));idx={p:i for i,p in enumerate(v)}
        f=[tuple(idx[vv[i]] for i in face) for face in ff]
        out,faces,c=split_contact_edges(v,f)
        self.assertEqual(c['contact_edges_split'],1)
        self.assertEqual(c['new_surface_faces'],0)
        self.assertEqual(triangle_hash(v,f),triangle_hash(out,faces))

    def test_hole_rejected(self):
        with self.assertRaisesRegex(TopologyError,'Open'):split_contact_edges(V,F[:-1])

    def test_reverse_rejected(self):
        with self.assertRaisesRegex(TopologyError,'winding'):split_contact_edges(V,[F[0][::-1]]+F[1:])

    def test_t_junction(self):
        v=V+[(.5,0,0)];f=[F[0],(0,4,3),(4,1,3),F[2],F[3]]
        out,faces,c=split_t_seams(v,f,1e-8)
        split_contact_edges(out,faces)
        self.assertEqual(c['new_caps'],0)
        self.assertEqual(c['maximum_bidirectional_surface_displacement_bound_mm'],0)

    def test_near_t_junction_bound_and_rejection(self):
        v=V+[(.5,.0001,0)];f=[F[0],(0,4,3),(4,1,3),F[2],F[3]]
        out,faces,c=split_t_seams(v,f,.0002)
        split_contact_edges(out,faces)
        self.assertAlmostEqual(c['maximum_bidirectional_surface_displacement_bound_mm'],.0001)
        out,faces,c=split_t_seams(v,f,.00001)
        with self.assertRaises(TopologyError):split_contact_edges(out,faces)

    def test_open_loop_never_capped(self):
        out,faces,c=split_t_seams(V,F[:-1],.1)
        self.assertEqual(len(faces),3)
        with self.assertRaises(TopologyError):split_contact_edges(out,faces)
