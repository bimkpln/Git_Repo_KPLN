import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location("pipe_slab_server",Path(__file__).resolve().parents[1]/"server"/"navis_mcp_server.py")
server = importlib.util.module_from_spec(spec)
spec.loader.exec_module(server)


def sample():
    return dict(Item1Path="Перекрытия",Item2Path="Трубы",Item2Size="ø100 мм",
                PipeSlabGeometry=dict(Usable=True, Source="pipe-slab-mesh-1", PipeSide=2,
                    PipeAxis=dict(X=0,Y=0,Z=1), SlabNormal=dict(X=0,Y=0,Z=-1),
                    OuterRadius=54/304.8, PipeId="p",SlabId="s", EdgeContact=False,
                    FullSectionCovered=True,Sections=[dict(Origin=dict(X=0,Y=0,Z=0),
                    Contour=[dict(X=1,Y=0,Z=0),dict(X=0,Y=1,Z=0),dict(X=-1,Y=0,Z=0)])]))


class SlabMetricsTests(unittest.TestCase):
    def test_normalization_and_swapped_participants(self):
        r=sample()
        for reverse in (False,True):
            if reverse:
                r.update(Item1Path="Трубы",Item2Path="Перекрытия",Item1Size="ø100 мм",Item2Size=None)
                r["PipeSlabGeometry"]["PipeSide"]=1
            m=server._metrics(r)
            self.assertTrue(m["GeomOk"])
            self.assertAlmostEqual(m["SlabOuterDiameterMm"],108)
            self.assertEqual(m["SectionMax"],100)
            self.assertEqual(m["SlabNormalAngle"],0)
            self.assertEqual(m["SlabSections"][0]["contour_mm"][0],[304.8,0,0])

    def test_old_bridge_does_not_fallback_to_aabb(self):
        r=sample();del r["PipeSlabGeometry"]
        r["Item1Bound"]=dict(Min=dict(X=0,Y=0,Z=0),Max=dict(X=100,Y=100,Z=1))
        r["Item2Bound"]=dict(Min=dict(X=0,Y=0,Z=0),Max=dict(X=1,Y=1,Z=10))
        m=server._metrics(r)
        self.assertFalse(m["GeomOk"]);self.assertIsNone(m["SlabNormalAngle"])
        self.assertEqual(m["SectionMax"],100)

    def test_invalid_axis_and_version_rejected(self):
        for key,value in (("PipeAxis",dict(X=0,Y=0,Z=0)),("Source","unknown"),("Usable",False),("OuterRadius",float("nan"))):
            r=sample();r["PipeSlabGeometry"][key]=value
            self.assertFalse(server._metrics(r)["GeomOk"])

    def test_missing_surface_fields_are_not_false(self):
        r=sample();del r["PipeSlabGeometry"]["EdgeContact"]
        self.assertIsNone(server._metrics(r)["SlabEdgeContact"])


if __name__ == "__main__":unittest.main()
