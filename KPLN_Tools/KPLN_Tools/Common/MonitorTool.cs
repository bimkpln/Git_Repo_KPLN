using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    internal class MonitorTool
    {
        internal static double? GetDoubleValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.Double:
                    return p.AsDouble();
                case StorageType.Integer:
                    return p.AsInteger();
                case StorageType.String:
                    return double.Parse(p.AsString(), System.Globalization.NumberStyles.Float);
                default:
                    return null;
            }
        }

        internal static int? GetIntegerValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.Double:
                    return (int)Math.Round(p.AsDouble());
                case StorageType.Integer:
                    return p.AsInteger();
                case StorageType.String:
                    return int.Parse(p.AsString(), System.Globalization.NumberStyles.Integer);
                default:
                    return null;
            }
        }

        internal static string GetStringValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.ElementId:
                        return p.AsValueString();
                    case StorageType.Double:
                        return p.AsValueString();
                    case StorageType.Integer:
                        return p.AsValueString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                switch (p.StorageType)
                {
                    case StorageType.ElementId:
                        return p.AsValueString();
                    case StorageType.Double:
                        return p.AsDouble().ToString();
                    case StorageType.Integer:
                        return p.AsInteger().ToString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
        }

        internal static Solid GetSolidFromElem(Element elem)
        {
            Solid resultSolid = null;
            GeometryElement geomElem = elem.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine });
            foreach (GeometryObject gObj in geomElem)
            {
                Solid solid = gObj as Solid;
                GeometryInstance gInst = gObj as GeometryInstance;
                if (solid != null) resultSolid = solid;
                else if (gInst != null)
                {
                    GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                    double tempVolume = 0;
                    foreach (GeometryObject gObj2 in instGeomElem)
                    {
                        solid = gObj2 as Solid;
                        if (solid != null && solid.Volume > tempVolume)
                        {
                            tempVolume = solid.Volume;
                            resultSolid = solid;
                        }
                    }
                }
            }

            return resultSolid;
        }

        internal static List<Curve> GetCurvesListFromPoints(List<XYZ> pointsOfIntersect)
        {
            List<Curve> curvesList = new List<Curve>();
            for (int i = 0; i < pointsOfIntersect.Count; i++)
            {
                if (i == pointsOfIntersect.Count - 1)
                {
                    curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[0]));
                    continue;
                }

                curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[i + 1]));
            }
            return curvesList;
        }
    }
}