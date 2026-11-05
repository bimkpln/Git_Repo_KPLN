"""Paint actual polygon material without erasing earlier layers in its holes."""
def paint_polygons(image,polygons,project,fill):
    from PIL import Image,ImageDraw,ImageChops
    mask=Image.new('L',image.size,0)
    for polygon in polygons:
        part=Image.new('L',image.size,0);draw=ImageDraw.Draw(part)
        draw.polygon(project(polygon.exterior.coords),fill=255)
        for ring in polygon.interiors:draw.polygon(project(ring.coords),fill=0)
        mask=ImageChops.lighter(mask,part)
    image.paste(fill,(0,0),mask)
