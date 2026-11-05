import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np
from shapely.geometry import box
from stair_finish import *
class FinishTests(unittest.TestCase):
    def test_without_finish_is_not_positive_evidence(self):
        self.assertFalse(explicit_finish({'Тип':'Монолитная без отделки 155х250'}))
    def test_explicit_finish_family(self):
        self.assertTrue(explicit_finish({'Тип':'Монолитная с отделкой 160х300_отделка плитка'}))
    def test_missing_semantics_not_guessed(self):self.assertFalse(explicit_finish({'Тип':'Монолитная 150х250'}))
    def test_translated_equal_thickness_not_excluded(self):
        z=lambda value:np.array([0.,0.,value])
        self.assertEqual(finish_region(box(0,0,10,10),z(170),z(20),z(150),z(0),20,0).area,0)
    def test_thickness_excess_and_matching_underside(self):
        z=lambda value:np.array([0.,0.,value])
        self.assertEqual(finish_region(box(0,0,10,10),z(170),z(0),z(150),z(0),20,.01).area,100)
    def test_misaligned_underface_remains(self):
        z=lambda value:np.array([0.,0.,value])
        self.assertEqual(finish_region(box(0,0,10,10),z(200),z(30),z(150),z(0),20,0).area,0)
    def test_only_positive_material_intervals(self):
        z=lambda value:np.array([0.,0.,value])
        self.assertEqual(finish_region(box(0,0,10,10),z(10),z(20),z(150),z(0),20,0).area,0)
if __name__=='__main__':unittest.main()
