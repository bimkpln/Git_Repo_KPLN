import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import manifold3d as md
from conditioning_topology import *
class ConditioningTopologyTests(unittest.TestCase):
    def test_rigid_transform_preserves_topology(self):
        box=md.Manifold.cube([10,10,10],center=True)
        self.assertTrue(require_same_topology(box,box.rotate([0,0,31]).translate([40000,50000,70000]))['passed'])
    def test_artificial_tunnel_rejected_even_if_mesh_valid(self):
        box=md.Manifold.cube([10,10,10],center=True);tube=md.Manifold.cylinder(20,2,2,center=True)
        with self.assertRaises(TopologyError):require_same_topology(box,box-tube)
    def test_deleted_tunnel_rejected(self):
        box=md.Manifold.cube([10,10,10],center=True);tube=md.Manifold.cylinder(20,2,2,center=True)
        with self.assertRaises(TopologyError):require_same_topology(box-tube,box)
    def test_disconnected_shell_count_preserved(self):
        box=md.Manifold.cube([10,10,10],center=True);two=box+box.translate([20,0,0])
        with self.assertRaises(TopologyError):require_same_topology(two,md.Manifold.cube([20,10,10],center=True))
if __name__=='__main__':unittest.main()
