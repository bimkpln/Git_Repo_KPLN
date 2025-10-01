using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckMEPHeight
{
    public sealed class CheckMEPHeightARElemData
    {
        private readonly List<Solid> _arSolids = new List<Solid>();
        private readonly FaceArray _arElementDownFacesArray = new FaceArray();
        private readonly List<BoundingBoxXYZ> _arBBoxes = new List<BoundingBoxXYZ>();
        private Transform _arElemLinkTrans;
        private static readonly object _locker = new object();

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
                    Options opt = new Options
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                        ComputeReferences = true
                    };
                    GeometryElement geomElem = ARElement.get_Geometry(opt);
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
                    Options opt = new Options
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                        ComputeReferences = true
                    };
                    GeometryElement geometryElement = ARElement.get_Geometry(opt);
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
        /// Коллекция ВЕРХНИХ поверхностей для проекции
        /// </summary>
        public FaceArray ARElemUpFacesArray
        {
            get
            {
                lock (_locker)
                {
                    if (_arElementDownFacesArray.IsEmpty)
                    {
                        foreach (Solid solid in ARElemSolids)
                        {
                            FaceArray solidFaceArray = GetHorizontalDownFacesFromArray(solid.Faces, false);
                            foreach (Face face in solidFaceArray)
                            {
                                _arElementDownFacesArray.Append(face);
                            }
                        }
                    }

                    return _arElementDownFacesArray;
                }
            }
        }

        /// <summary>
        /// Коллекция НИЖНИХ поверхностей для проекции 
        /// </summary>
        public FaceArray ARElemDownFacesArray
        {
            get
            {
                lock (_locker)
                {
                    if (_arElementDownFacesArray.IsEmpty)
                    {
                        foreach (Solid solid in ARElemSolids)
                        {
                            FaceArray solidFaceArray = GetHorizontalDownFacesFromArray(solid.Faces, true);
                            foreach (Face face in solidFaceArray)
                            {
                                _arElementDownFacesArray.Append(face);
                            }
                        }
                    }

                    return _arElementDownFacesArray;
                }
            }
        }

        /// <summary>
        /// Получить FaceArray ТОЛЬКО горизонтальных нижних/верхних плоскостей
        /// </summary>
        /// <param name="faceArray"></param>
        /// <returns></returns>
        public static FaceArray GetHorizontalDownFacesFromArray(FaceArray faceArray, bool isDownFace)
        {
            FaceArray result = new FaceArray();
            foreach (Face face in faceArray)
            {
                if (isDownFace)
                {
                    // Фильтрация PlanarFace, которые являются боковыми или нижними гранями
                    if (face is PlanarFace planarFace && (Math.Abs(planarFace.FaceNormal.X) > 0.1 || Math.Abs(planarFace.FaceNormal.Y) > 0.1 || planarFace.FaceNormal.Z < 0))
                        continue;

                    // Фильтрация CylindricalFace, которые являются боковыми или нижними гранями
                    if (face is CylindricalFace cylindricalFace && (Math.Abs(cylindricalFace.Axis.X) > 0.1 || Math.Abs(cylindricalFace.Axis.Y) > 0.1 || cylindricalFace.Axis.Z < 0))
                        continue;
                }
                else
                {
                    // Фильтрация PlanarFace, которые являются боковыми или нижними гранями
                    if (face is PlanarFace planarFace && (Math.Abs(planarFace.FaceNormal.X) > 0.1 || Math.Abs(planarFace.FaceNormal.Y) > 0.1 || planarFace.FaceNormal.Z > 0))
                        continue;

                    // Фильтрация CylindricalFace, которые являются боковыми или нижними гранями
                    if (face is CylindricalFace cylindricalFace && (Math.Abs(cylindricalFace.Axis.X) > 0.1 || Math.Abs(cylindricalFace.Axis.Y) > 0.1 || cylindricalFace.Axis.Z > 0))
                        continue;
                }

                result.Append(face);
            }

            return result;
        }

        public override bool Equals(object obj)
        {
            if (obj is CheckMEPHeightARElemData item)
                return this.ARElement.Id.IntegerValue == item.ARElement.Id.IntegerValue;

            return false;
        }

        public override int GetHashCode() => this.ARElement.Id.GetHashCode();

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
