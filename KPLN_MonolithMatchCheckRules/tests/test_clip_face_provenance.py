import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from clip_face_provenance import classify_faces

class ClipProvenanceTests(unittest.TestCase):
    def test_real_source_face_on_clip_plane_is_retained(self):
        tri=[[[0,1,0],[1,1,0],[0,1,1]]];owners,caps=classify_faces(tri,[7],{7:12},{9},1,1e-6)
        self.assertEqual(owners.tolist(),[12]);self.assertFalse(caps[0])
    def test_known_numerical_clip_face_is_identified(self):
        tri=[[[0,1+1e-8,0],[1,1,0],[0,1,1]]];_,caps=classify_faces(tri,[9],{},[9],1,1e-6)
        self.assertTrue(caps[0])
    def test_clipping_label_never_excuses_off_plane_vertex(self):
        with self.assertRaises(ValueError):classify_faces([[[0,1,0],[1,1,0],[0,.8,0]]],[9],{},[9],1,1e-6)
    def test_unknown_and_colliding_identities_rejected(self):
        tri=[[[0,1,0],[1,1,0],[0,1,1]]]
        with self.assertRaises(ValueError):classify_faces(tri,[3],{},[9],1,1e-6)
        with self.assertRaises(ValueError):classify_faces(tri,[7],{7:0},[7],1,1e-6)
if __name__=='__main__':unittest.main()
