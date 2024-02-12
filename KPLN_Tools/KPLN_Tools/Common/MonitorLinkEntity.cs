using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    public class MonitorLinkEntity
    {
        private Solid _linkElementSolid;

        public MonitorLinkEntity(Element linkElement, HashSet<Parameter> linkElemsParams, RevitLinkInstance linkIsntance)
        {
            LinkElement = linkElement;
            LinkIsntance = linkIsntance;
            LinkElemsParams = linkElemsParams;
        }
        
        public Element LinkElement { get; protected set; }

        public Solid LinkElementSolid
        {
            get
            {
                try
                {
                    if (_linkElementSolid == null)
                    {
                        Solid resultSolid = MonitorTool.GetSolidFromElem(LinkElement);
                        _linkElementSolid = LinkIsntance == null ? resultSolid : SolidUtils.CreateTransformed(resultSolid, LinkIsntance.GetTotalTransform());
                    }
                }
                catch (Exception)
                {
                    // Случай, когда у эл-та нет геометрии (редкая ошибка, которая не должна влиять на смежника)
                    LocationPoint locPoint = LinkElement.Location as LocationPoint;
                    XYZ xyzPnt = locPoint.Point;
                    List<XYZ> pointsDwn = new List<XYZ>()
                    {
                        new XYZ(xyzPnt.X - 0.5, xyzPnt.Y - 0.5, xyzPnt.Z - 0.5),
                        new XYZ(xyzPnt.X - 0.5, xyzPnt.Y + 0.5, xyzPnt.Z - 0.5),
                        new XYZ(xyzPnt.X + 0.5, xyzPnt.Y + 0.5, xyzPnt.Z - 0.5),
                        new XYZ(xyzPnt.X + 0.5, xyzPnt.Y - 0.5, xyzPnt.Z - 0.5),
                    };
                    List<XYZ> pointsUp = new List<XYZ>()
                    {
                        new XYZ(xyzPnt.X - 0.5, xyzPnt.Y - 0.5, xyzPnt.Z + 0.5),
                        new XYZ(xyzPnt.X - 0.5, xyzPnt.Y + 0.5, xyzPnt.Z + 0.5),
                        new XYZ(xyzPnt.X + 0.5, xyzPnt.Y + 0.5, xyzPnt.Z + 0.5),
                        new XYZ(xyzPnt.X + 0.5, xyzPnt.Y - 0.5, xyzPnt.Z + 0.5),
                    };

                    List<Curve> curvesListDwn = MonitorTool.GetCurvesListFromPoints(pointsDwn);
                    List<Curve> curvesListUp = MonitorTool.GetCurvesListFromPoints(pointsUp);
                    CurveLoop curveLoopDwn = CurveLoop.Create(curvesListDwn);
                    CurveLoop curveLoopUp = CurveLoop.Create(curvesListUp);

                    CurveLoop[] curves = new CurveLoop[] { curveLoopDwn, curveLoopUp };

                    SolidOptions solidOptions = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);

                    Solid linkSOlid = GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions);
                    _linkElementSolid = SolidUtils.CreateTransformed(linkSOlid, LinkIsntance.GetTotalTransform());

                }

                return _linkElementSolid;
            }
        }
        public RevitLinkInstance LinkIsntance { get; protected set; }
        public HashSet<Parameter> LinkElemsParams { get; protected set; }
    }
}
