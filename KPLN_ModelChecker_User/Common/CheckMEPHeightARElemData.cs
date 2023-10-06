using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_User.Common
{
    internal class CheckMEPHeightARElemData
    {
        private readonly List<Solid> _arSolids = new List<Solid>();
        private readonly FaceArray _arElementDownFacesArray = new FaceArray();
        private readonly List<BoundingBoxXYZ> _arBBoxes = new List<BoundingBoxXYZ>();
        private Transform _arElemLinkTrans;

        public CheckMEPHeightARElemData(Element arElemen, RevitLinkInstance arElementLinkInst)
        {
            ARElement = arElemen;
            ARElementLinkInst = arElementLinkInst;
        }

        public Element ARElement { get; set; }

        public RevitLinkInstance ARElementLinkInst { get; set; }

        public Transform ARElemLinkTrans
        {
            get
            {
                if (_arElemLinkTrans == null)
                    _arElemLinkTrans = ARElementLinkInst.GetTotalTransform();

                return _arElemLinkTrans;
            }
        }

        /// <summary>
        /// Коллекция солидов у элемента АР
        /// </summary>
        public List<Solid> ARElemSolids
        {
            get
            {
                if (_arSolids.Count == 0)
                {
                    GeometryElement geomElem = ARElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });
                    if (geomElem != null)
                    {
                        GetSolidsFromGeomElem(geomElem, Transform.Identity, _arSolids);
                    }

                    if (_arSolids == null || _arSolids.Count == 0)
                        throw new Exception($"Не удалось получить полноценную коллекцию Solid у элемента с id: {ARElement.Id}");
                }

                return _arSolids;
            }
        }

        /// <summary>
        /// BoundingBoxXYZ у элемента АР
        /// </summary>
        public List<BoundingBoxXYZ> ARElemBBoxes
        {
            get
            {
                if (_arBBoxes.Count == 0)
                {
                    GeometryElement geometryElement = ARElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });
                    foreach (GeometryObject geomObject in geometryElement)
                    {
                        switch (geomObject)
                        {
                            case Solid solid:
                                _arBBoxes.Add(GetBoundingBoxXYZ(solid));
                                break;
                            case GeometryInstance geomInstance:
                                GeometryElement instGeomElem = geomInstance.GetInstanceGeometry();
                                _arBBoxes.AddRange(GetBoundingBoxXYZColl(instGeomElem));
                                break;

                            case GeometryElement geomElem:
                                _arBBoxes.AddRange(GetBoundingBoxXYZColl(geomElem));
                                break;
                        }
                        if (_arBBoxes.Count == 0)
                            throw new Exception($"Не удалось получить BoundingBoxXYZ у элемента с id: {ARElement.Id}");
                    }
                }

                return _arBBoxes;
            }
        }

        /// <summary>
        /// Коллекция поверхностей для проекции
        /// </summary>
        public FaceArray ARElemDownFacesArray 
        {
            get
            {
                if (_arElementDownFacesArray.IsEmpty)
                {
                    foreach(Solid solid in ARElemSolids)
                    {
                        FaceArray faceArray = solid.Faces;
                        foreach (Face face in faceArray)
                        {
                            // Фильтрация PlanarFace, которые являются боковыми гранями
                            if (face is PlanarFace planarFace && (Math.Abs(planarFace.FaceNormal.X) > 0.1 || Math.Abs(planarFace.FaceNormal.Y) > 0.1))
                                continue;

                            // Фильтрация CylindricalFace, которые являются боковыми гранями
                            if (face is CylindricalFace cylindricalFace && (Math.Abs(cylindricalFace.Axis.X) > 0.1 || Math.Abs(cylindricalFace.Axis.Y) > 0.1))
                                continue;

                            _arElementDownFacesArray.Append(face);
                        }
                    }
                }

                return _arElementDownFacesArray;
            }
        }

        /// <summary>
        /// Получить солид из элементов АР
        /// </summary>
        private void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solids)
        {
            foreach (GeometryObject geomObject in geometryElement)
            {
                switch (geomObject)
                {
                    case Solid solid:
                        if (solid.Volume > 0) solids.Add(solid);
                        break;

                    case GeometryInstance geomInstance:
                        GetSolidsFromGeomElem(geomInstance.GetInstanceGeometry(), geomInstance.Transform.Multiply(transformation), solids);
                        break;

                    case GeometryElement geomElem:
                        GetSolidsFromGeomElem(geomElem, transformation, solids);
                        break;
                }
            }
        }

        private List<BoundingBoxXYZ> GetBoundingBoxXYZColl(GeometryElement geomElem)
        {
            List<BoundingBoxXYZ> result = new List<BoundingBoxXYZ>();
            foreach (GeometryObject obj in geomElem)
            {
                Solid solid = obj as Solid;
                BoundingBoxXYZ bbox = GetBoundingBoxXYZ(solid);
                if (bbox != null)
                {
                    result.Add(bbox);
                }
            }

            return result;
        }


        private BoundingBoxXYZ GetBoundingBoxXYZ(Solid solid)
        {
            if (solid != null && solid.Volume != 0)
            {
                BoundingBoxXYZ bbox = solid.GetBoundingBox();
                Transform transform = bbox.Transform;
                Transform resultTrans = transform * ARElemLinkTrans;
                return new BoundingBoxXYZ()
                {
                    Max = resultTrans.OfPoint(bbox.Max),
                    Min = resultTrans.OfPoint(bbox.Min),
                };
            }

            return null;
        }
    }
}
