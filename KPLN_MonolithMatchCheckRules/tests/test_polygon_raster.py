import unittest
from PIL import Image
from shapely.geometry import Polygon,box
from polygon_raster import paint_polygons

class PolygonRasterTests(unittest.TestCase):
    def test_hole_keeps_existing_common_material(self):
        image=Image.new('RGB',(12,12),'gray')
        ring=Polygon([(1,1),(10,1),(10,10),(1,10)],holes=[[(3,3),(8,3),(8,8),(3,8)]])
        paint_polygons(image,[ring],lambda p:list(p),'red')
        self.assertEqual(image.getpixel((5,5)),(128,128,128))
        self.assertEqual(image.getpixel((2,2)),(255,0,0))
        self.assertEqual(image.getpixel((0,0)),(128,128,128))
    def test_separate_island_does_not_depend_on_polygon_order(self):
        outer=Polygon([(0,0),(11,0),(11,11),(0,11)],holes=[[(2,2),(9,2),(9,9),(2,9)]])
        island=box(4,4,7,7)
        images=[]
        for parts in [[island,outer],[outer,island]]:
            image=Image.new('RGB',(12,12),'gray');paint_polygons(image,parts,lambda p:list(p),'green');images.append(image)
        self.assertEqual(images[0].tobytes(),images[1].tobytes())
        self.assertEqual(images[0].getpixel((5,5)),(0,128,0))
        self.assertEqual(images[0].getpixel((3,3)),(128,128,128))
    def test_empty_layer_preserves_whole_image(self):
        image=Image.new('RGB',(5,5),'gray');before=image.tobytes()
        paint_polygons(image,[],lambda p:list(p),'red');self.assertEqual(image.tobytes(),before)

if __name__=='__main__':unittest.main()
