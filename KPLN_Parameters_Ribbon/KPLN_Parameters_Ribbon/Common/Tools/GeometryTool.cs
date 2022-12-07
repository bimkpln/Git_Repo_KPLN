using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    internal static class GeometryTool
    {
        public static XYZ[] GetMaxMinHeightPoints(List<Solid> solids)
        {
            XYZ maxZpoint = new XYZ(0, 0, -9999999);
            XYZ minZpount = new XYZ(0, 0, 9999999);

            List<Edge> edges = new List<Edge>();
            foreach (Solid s in solids)
            {
                foreach (Edge e in s.Edges)
                {
                    edges.Add(e);
                }
            }

            foreach (Edge e in edges)
            {
                Curve c = e.AsCurve();
                XYZ p1 = c.GetEndPoint(0);
                if (p1.Z > maxZpoint.Z) maxZpoint = p1;
                if (p1.Z < minZpount.Z) minZpount = p1;

                XYZ p2 = c.GetEndPoint(1);
                if (p2.Z > maxZpoint.Z) maxZpoint = p2;
                if (p2.Z < minZpount.Z) minZpount = p2;
            }
            XYZ[] result = new XYZ[] { maxZpoint, minZpount };
            return result;
        }


        public static List<Solid> GetSolidsFromElement(Element elem)
        {
            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;
            GeometryElement geoElem = elem.get_Geometry(opt);

            List<Solid> solids = GetSolidsFromElement(geoElem);
            return solids;
        }

        public static List<Solid> GetSolidsFromElement(GeometryElement geoElem)
        {
            List<Solid> solids = new List<Solid>();

            if (geoElem != null)
            {
                foreach (GeometryObject geomObj in geoElem)
                {
                    if (geomObj is Solid)
                    {
                        solids.Add((Solid)geomObj);
                    }
                    else if (geomObj is GeometryInstance geomInst)
                    {
                        if (geomInst.GetInstanceGeometry() is GeometryElement instGeomEl)
                        {
                            foreach (var instGeomObj in instGeomEl)
                            {
                                if (instGeomObj is Solid instanceGeometrySolid)
                                {
                                    solids.Add(instanceGeometrySolid);
                                }
                            }
                        }
                    }
                }
            }
            return solids;
        }

    }
}
