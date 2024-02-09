using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    public class MonitorEntity
    {
        private Solid _linkElementSolid;

        internal MonitorEntity(Element linkElement, RevitLinkInstance linkIsntance)
        {
            LinkElement = linkElement;
            LinkIsntance = linkIsntance;
        }

        internal Element LinkElement { get; }
        internal Solid LinkElementSolid
        {
            get
            {
                if (_linkElementSolid == null)
                {
                    Solid resultSolid = GetSolidFromElem(LinkElement);
                    _linkElementSolid = LinkIsntance == null ? resultSolid : SolidUtils.CreateTransformed(resultSolid, LinkIsntance.GetTotalTransform());
                }

                return _linkElementSolid;
            }
        }
        internal RevitLinkInstance LinkIsntance { get; }
        internal HashSet<Parameter> LinkElemsParams { get; set; }

        internal Element ModelElement { get; set; }
        internal Solid ModelElementSolid { get; set; }
        internal HashSet<Parameter> ModelParameters { get; set; }

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
    }
}
