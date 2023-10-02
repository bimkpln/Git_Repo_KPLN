using Autodesk.Revit.DB;
using System;

namespace KPLN_ModelChecker_User.Common
{
    internal class CheckMEPHeightARElemData
    {
        private Solid _arSolid;
        private Transform _arLinkTrans;

        public CheckMEPHeightARElemData(Element arElemen, RevitLinkInstance arElementLinkInst)
        {
            ARElement = arElemen;
            ARElementLinkInst = arElementLinkInst;
        }

        public Element ARElement { get; set; }

        public RevitLinkInstance ARElementLinkInst { get; set; }

        public Transform ARLinkTrans
        {
            get
            {
                if (_arLinkTrans == null)
                    _arLinkTrans = ARElementLinkInst.GetTotalTransform();

                return _arLinkTrans;
            }
        }

        public Solid ARSolid
        {
            get
            {
                if (_arSolid == null)
                {
                    GeometryElement geomElem = ARElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });

                    foreach (GeometryObject gObj in geomElem)
                    {
                        if (gObj is Solid solid1)
                        {
                            _arSolid = solid1;
                            break;
                        }
                        else if (gObj is GeometryInstance gInst)
                        {
                            GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                            double tempVolume = 0;
                            foreach (GeometryObject gObj2 in instGeomElem)
                            {
                                if (gObj2 is Solid solid2 && solid2.Volume > tempVolume)
                                    _arSolid = solid2;
                            }
                        }
                    }

                    if (_arSolid == null)
                        throw new Exception($"Не удалось получить геометрию у элемента с id: {ARElement.Id}");
                }

                return _arSolid;
            }
        }

        /// <summary>
        /// Коллекция поверхностей для проекции
        /// </summary>
        public FaceArray ARElementDownFacesArray { get; set; } = new FaceArray();
    }
}
