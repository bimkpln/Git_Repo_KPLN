import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import unittest
import numpy as np,trimesh,shapely
from shapely.geometry import box
from triangle_distance_masks import *

class AnalyticDistanceMaskTests(unittest.TestCase):
    def test_near_duplicate_oblique_triangle_neighbourhoods(self):
        v=[[-20,-10,-5],[20,-10,8],[0,20,3],[-20,-10,-5+1e-10],[20,-10,8+1e-10],[0,20,3+1e-10],[20,20,3]]
        self.verify_points(v,[[0,1,2],[3,4,5],[1,6,2]])
    def test_two_parallel_boundaries_below_thirty_cover_vertical_midgap(self):
        gap=29.999985694885254
        vertices=np.array([[-100,-100,0],[300,-100,0],[-100,300,0],[-100,-100,gap],[300,-100,gap],[-100,300,gap]])
        B=np.array([[1.,0,0],[0,0,1.],[0,-1.,0]])
        roi=box(-10,14.9999,10,15.0001)
        inside,over,unknown,p=TriangleDistanceMasks(vertices,[[0,1,2],[3,4,5]]).masks(np.zeros(3),B,roi,15-2e-7,15+2e-7,guard_mm=1e-7)
        self.assertTrue(inside.covers(roi));self.assertEqual(over.area,0);self.assertEqual(unknown.area,0)

    def test_overlapping_coplanar_faces_do_not_cancel_existing_support(self):
        import json
        from pathlib import Path
        data=json.loads(Path(__file__).with_name('mask_coplanar_overlap_fixture.json').read_text())
        inside,over,unknown,proof=TriangleDistanceMasks(data['vertices'],data['faces']).masks(
            np.zeros(3),np.eye(3),shapely.from_wkt(data['roi']),data['within_radius'],data['over_radius'])
        point=Point(data['point'])
        self.assertTrue(inside.covers(point))
        self.assertFalse(over.covers(point))

    def verify_points(self,vertices,faces,radius=3):
        roi=box(-30,-30,30,30);v=np.asarray(vertices,float);f=np.asarray(faces,int)
        inside,over,unknown,proof=TriangleDistanceMasks(v,f).masks(np.zeros(3),np.eye(3),roi,radius,radius)
        points=np.random.default_rng(37).uniform(-29.9,29.9,(5000,2));p3=np.c_[points,np.zeros(len(points))]
        mesh=trimesh.Trimesh(vertices=v,faces=f,process=False)
        _,dist,_=trimesh.proximity.closest_point_naive(mesh,p3)
        shapes=shapely.points(points);mask_inside=shapely.contains(inside,shapes);mask_over=shapely.contains(over,shapes)
        self.assertTrue(np.all(dist[mask_inside]<=radius+1e-6))
        self.assertTrue(np.all(dist[mask_over]>=radius-1e-6))
        self.assertAlmostEqual(inside.area+over.area+unknown.area,roi.area,places=7)
        self.assertLess(inside.intersection(over).area,1e-8)

    def test_oblique_face_prism_and_edge_ellipses(self):
        self.verify_points([[-20,-10,-5],[20,-10,8],[0,20,3]],[[0,1,2]])

    def test_perpendicular_triangle(self):
        self.verify_points([[-20,0,-10],[20,0,-10],[0,0,20]],[[0,1,2]])

    def test_parallel_plane_has_correct_height_radius(self):
        self.verify_points([[0,0,10],[10,0,10],[0,10,10]],[[0,1,2]],radius=15)

    def test_vertical_degenerate_edge_cylinder(self):
        roi=box(-5,-5,5,5);v=np.array([[0,0,-10],[0,0,10]],float)
        a,b,u,p=TriangleDistanceMasks(v,[[0,1,1]]).masks(np.zeros(3),np.eye(3),roi,3,3)
        self.assertLessEqual(a.area,np.pi*9)
        self.assertGreaterEqual(roi.area-b.area,np.pi*9)
        self.assertLess(u.area,.1)

    def test_almost_parallel_degenerate_edge(self):
        roi=box(-30,-30,30,30);v=np.array([[-20,0,2],[20,0,2+1e-9]],float)
        a,b,u,p=TriangleDistanceMasks(v,[[0,1,1]]).masks(np.zeros(3),np.eye(3),roi,3,3)
        self.assertTrue(a.covers(Point(0,1)))
        self.assertTrue(b.covers(Point(0,5)))

    def test_exact_threshold_not_an_issue(self):
        v=np.array([[-100,-100,15],[300,-100,15],[-100,300,15]],float)
        a,b,u,p=TriangleDistanceMasks(v,[[0,1,2]]).masks(np.zeros(3),np.eye(3),box(0,0,10,10),15,15,guard_mm=0)
        self.assertAlmostEqual(a.area,100)
        self.assertEqual(b.area,0)

    def test_empty_target_has_no_fabricated_surface(self):
        a,b,u,p=TriangleDistanceMasks(np.empty((0,3)),np.empty((0,3),int)).masks(np.zeros(3),np.eye(3),box(0,0,10,10),15,15)
        self.assertEqual(a.area,0);self.assertEqual(b.area,100);self.assertEqual(u.area,0)

    def test_near_duplicate_ring_vertex_never_loses_a_finite_triangle(self):
        polygon=Polygon([(190,-14.980524027500039),(190,-14.98052402750004),(190,-76.3859843400682),
            (190,-935.0194759724991),(190,-949.9999999999991),(189.99999999999997,-949.9999999999991),
            (175.01947597249995,-949.9999999999991),(93.94795562864705,-949.9999999999991),
            (0,-949.9999999999991),(190,0)])
        triangles,proof=certificate_triangulation(polygon)
        self.assertLess(polygon.symmetric_difference(shapely.union_all(triangles)).area,1e-5)
        self.assertLessEqual(proof['normalization_coordinate_bound_mm'],1e-9)

if __name__=='__main__':unittest.main()

