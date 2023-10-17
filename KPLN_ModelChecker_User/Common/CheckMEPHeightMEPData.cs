using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по каждому элементу ИОС
    /// </summary>
    internal class CheckMEPHeightMEPData
    {
        /// <summary>
        /// От данного парамтера зависит точность опредления помещений для линейных элементов
        /// </summary>
        private const int _bboxArrayCount = 5;
        /// <summary>
        /// От данного парамтера зависит точность опредления привязки элемента к помещению
        /// </summary>
        private const int _bboxExpanded = 10;
        private readonly List<BoundingBoxXYZ> _mepElemBBoxes = new List<BoundingBoxXYZ>();
        public List<Solid> _mepElemSolids = new List<Solid>();
        private BoundingBoxXYZ[] _currentBBoxArray;
        /// <summary>
        /// Точность в проверке при анализе положения элементов
        /// </summary>
        private static readonly double _toleranceToCheck = 0.1;

        public CheckMEPHeightMEPData(Element elem)
        {
            MEPElement = elem;
        }

        public Element MEPElement { get; }

        public List<Solid> MEPElemSolids
        {
            get
            {
                if (_mepElemSolids.Count == 0)
                {

                    Options opt = new Options() { DetailLevel = ViewDetailLevel.Fine };
                    opt.ComputeReferences = true;
                    GeometryElement geomElem = MEPElement.get_Geometry(opt);
                    if (geomElem != null)
                    {
                        GetSolidsFromGeomElem(geomElem, Transform.Identity, _mepElemSolids);
                    }

                    // Нужно отфильтровать на безсолидные элементы (изоляция отводов, элементы без геометрии и т.п.)
                    //if (_mepElemSolids == null || _mepElemSolids.Count == 0)
                    //    throw new Exception($"Не удалось получить полноценную коллекцию Solid у элемента с id: {MEPElement.Id}");
                }

                return _mepElemSolids;
            }
        }

        public List<BoundingBoxXYZ> MEPElemBBoxes
        {
            get
            {
                if (_mepElemBBoxes.Count == 0)
                {
                    GeometryElement geometryElement = MEPElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });
                    foreach (GeometryObject geomObject in geometryElement)
                    {
                        switch (geomObject)
                        {
                            case Solid solid:
                                _mepElemBBoxes.Add(GetBoundingBoxXYZ(solid));
                                break;
                            case GeometryInstance geomInstance:
                                GeometryElement instGeomElem = geomInstance.GetInstanceGeometry();
                                _mepElemBBoxes.AddRange(GetBoundingBoxXYZColl(instGeomElem));
                                break;

                            case GeometryElement geomElem:
                                _mepElemBBoxes.AddRange(GetBoundingBoxXYZColl(geomElem));
                                break;
                        }
                        if (_mepElemBBoxes.Count == 0)
                            throw new Exception($"Не удалось получить BoundingBoxXYZ у элемента с id: {MEPElement.Id}");
                    }
                }

                return _mepElemBBoxes;
            }
        }

        /// <summary>
        /// Массив равнозначных в плоскости XY BoundingBoxXYZ (для проверки по всем типам помещений)
        /// </summary>
        public BoundingBoxXYZ[] SplitedBBoxArray
        {
            get
            {
                if (_currentBBoxArray == null)
                {
                    if (MEPElement.Location is LocationCurve _) _currentBBoxArray = SetCurrentBBoxArray(_bboxArrayCount * MEPElemBBoxes.Count);
                    else _currentBBoxArray = MEPElemBBoxes.ToArray();
                }

                return _currentBBoxArray;
            }
        }

        /// <summary>
        /// Отсеиваю вертикальные линейные опуски или вертикальные транзиты ПО списку всех ошибок, имеющихся на данном этапе
        /// </summary>
        /// <param name="elemToCheck">Элемент для анализа</param>
        /// <param name="finallyCheckColl">Коллекция для поиска соседа</param>
        public static bool VerticalCurveElementsFilteredWithTolerance(Element elemToCheck, IEnumerable<CheckMEPHeightMEPData> finallyCheckColl)
        {
            if (elemToCheck.Location is LocationCurve locCurve)
            {
                Line locLine = locCurve.Curve as Line;
                XYZ direction = locLine.Direction;
                if (Math.Abs(direction.X) < _toleranceToCheck && Math.Abs(direction.Y) < _toleranceToCheck)
                {
                    if (elemToCheck is MEPCurve mepCurve)
                    {
                        ConnectorManager conManagerToCheck = mepCurve.ConnectorManager;
                        foreach (Connector connToCheck in conManagerToCheck.Connectors)
                        {
                            if (connToCheck.Domain == Domain.DomainUndefined) continue;

                            XYZ origin = connToCheck.Origin;
                            int conCount = 2;
                            foreach (CheckMEPHeightMEPData mepData in finallyCheckColl)
                            {
                                if (mepData.MEPElement.Id.IntegerValue == elemToCheck.Id.IntegerValue)
                                    continue;

                                if (mepData.MEPElement is FamilyInstance famInst)
                                {
                                    MEPModel mepModel = famInst.MEPModel;
                                    ConnectorManager conManager = mepModel.ConnectorManager;
                                    if (conManager != null)
                                    {
                                        foreach (Connector connector in conManager.Connectors)
                                        {
                                            XYZ famInstOrigin = connector.Origin;
                                            double distance = origin.DistanceTo(famInstOrigin);
                                            if (distance < _toleranceToCheck)
                                            {
                                                conCount--;
                                                if (conCount == 0)
                                                    // Проверка пройдена: Элемент линеен, и ограничен 2мя соседями, которые уже находятся в отчете
                                                    return true;
                                            }

                                        }
                                    }
                                }
                            }

                            // Проверка провалена: Элемент не имеет достаточное количество соседей, которые уже включены в отчет.
                            return false;
                        }

                    }
                    else
                        throw new Exception($"Не удалось проанализировать Location элемента ИОС с id: {elemToCheck.Id}");
                }
            }

            // Проверка пройдена: Элемент не линеен, либо не вертикален
            return true;
        }

        /// <summary>
        /// Определение находиться ли элемент в границах помещения
        /// </summary>
        /// <param name="arData">Спец. класс для проверки</param>
        public bool IsElemInCurrentRoom(CheckMEPHeightARRoomData arData)
        {
            foreach (BoundingBoxXYZ bbox in SplitedBBoxArray)
            {
                if (MEPElement.Id.IntegerValue == 13064543)
                {
                    var a = 1;
                }

                // Быстрая и неточная проверка на BoundingBoxXYZ
                if ((arData.RoomBBox.Max.X >= bbox.Min.X && arData.RoomBBox.Min.X <= bbox.Max.X)
                    && (arData.RoomBBox.Max.Y >= bbox.Min.Y && arData.RoomBBox.Min.Y <= bbox.Max.Y)
                    && (arData.RoomBBox.Max.Z + _bboxExpanded >= bbox.Min.Z && arData.RoomBBox.Min.Z <= bbox.Max.Z))
                {
                    // Более точная и длительная проверка на вхождение элемента в помещение (или над помещением, для органиченных помещений)
                    if (arData.CurrentRoom.IsPointInRoom(bbox.Min) 
                        || arData.CurrentRoom.IsPointInRoom(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z - _bboxExpanded / 5))
                        || arData.CurrentRoom.IsPointInRoom(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z - _bboxExpanded / 2))
                        || arData.CurrentRoom.IsPointInRoom(bbox.Max)
                        || arData.CurrentRoom.IsPointInRoom(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + _bboxExpanded / 5))
                        || arData.CurrentRoom.IsPointInRoom(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + _bboxExpanded / 2)))
                    {
                        return true;
                    }
                    
                    //// Более точная и длительная проверка на вхождение элемента в помещение
                    //foreach(Solid mepSolid in MEPElemSolids)
                    //{
                    //    EdgeArray edgeArray = mepSolid.Edges;
                    //    foreach(Edge edge in edgeArray)
                    //    {
                    //        Curve curve = edge.AsCurve();
                    //        if (curve != null)
                    //        {
                    //            SolidCurveIntersection intRes = arData
                    //                .RoomSolid
                    //                .IntersectWithCurve(curve, new SolidCurveIntersectionOptions(){ ResultType = SolidCurveIntersectionMode.CurveSegmentsInside });
                    //            int segmCount = intRes.SegmentCount;
                    //            if (intRes != null && segmCount > 0)
                    //                return true;
                    //        }
                    //    }
                    //}
                }
            }

            return false;
        }

        /// <summary>
        /// Проверить коллекцию эл-в ИОС нарушение высоты в помещении АР 
        /// </summary>
        /// <param name="currentRoomMEPDataColl">Коллекция спец. классов ИОС для проверки</param>
        /// <param name="arData">Спец. класс АР для проверки</param>
        public static CheckMEPHeightMEPData[] CheckIOSElemsForMinDistErrorByAR(CheckMEPHeightMEPData[] currentRoomMEPDataColl, CheckMEPHeightARRoomData arData) =>
            currentRoomMEPDataColl.Where(m => m.IsHeigtError(arData)).ToArray();

        /// <summary>
        /// Получить солид из элементов
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
                Transform resultTrans = transform;
                return new BoundingBoxXYZ()
                {
                    Max = resultTrans.OfPoint(bbox.Max),
                    Min = resultTrans.OfPoint(bbox.Min),
                };
            }

            return null;
        }


        /// <summary>
        /// Проверить элемент на факт нарушения высоты
        /// </summary>
        /// <param name="arData">Спец. класс для проверки</param>
        private bool IsHeigtError(CheckMEPHeightARRoomData arData)
        {
            if(MEPElement.Id.IntegerValue== 13064543)
            {
                var a = 1;
            }
            
            // Зада. мин длину, при которой элемент считается с потенциальным нарушением - выше 1.5 м
            double minIntDist = 5;
            double tempIntDist = Double.MaxValue;
            foreach (BoundingBoxXYZ checkBBox in SplitedBBoxArray)
            {
                #region Проверка вертикальных участков на вертикальную проходку между этажами (стояк). К ним отношу все участки, которые больше 1.5 м
                if (MEPElement.Location is LocationCurve locCurve)
                {
                    if (locCurve.Curve is Line locLine)
                    {
                        XYZ direction = locLine.Direction;
                        if (Math.Abs(direction.X) < _toleranceToCheck && Math.Abs(direction.Y) < _toleranceToCheck && locCurve.Curve.ApproximateLength > 5)
                            return false;
                    }
                    else
                        throw new Exception($"Не удалось проанализировать Location элемента ИОС с id: {MEPElement.Id}");
                }
                #endregion

                #region Проверка на минимальную дистанцию до поверхности
                XYZ[] pointsToCheck = new XYZ[5]
                {
                    GetCurrentBBoxZMin_HRight(checkBBox),
                    GetCurrentBBoxZMin_HLeft(checkBBox),
                    GetCurrentBBoxZMin_Center(checkBBox),
                    GetCurrentBBoxZMin_LRight(checkBBox),
                    GetCurrentBBoxZMin_LLeft(checkBBox),
                };

                foreach (XYZ point in pointsToCheck)
                {
                    foreach(CheckMEPHeightARElemData arElemData in arData.RoomDownARElemDataColl)
                    {
                        foreach (Face face in arElemData.ARElemDownFacesArray)
                        {
                            // Делаю инверсию точки элемента ИОС на координаты АР. Плоскость не подвергается трансформации координат, или созданию (нет конструктора)
                            XYZ inversedPoint = arElemData.ARElemLinkTrans.Inverse.OfPoint(point);
                            IntersectionResult intRes = face.Project(inversedPoint);
                            if (intRes != null && intRes.Distance < tempIntDist && intRes.Distance > minIntDist)
                            {
                                tempIntDist = intRes.Distance;
                                double iosDistance = inversedPoint.DistanceTo(intRes.XYZPoint);
                                if (iosDistance < arData.RoomMinDistance)
                                    return true;
                            }
                        }
                    }
                }
                #endregion
            }

            return false;
        }

        /// <summary>
        /// Получить правую верхнюю точку BoundingBoxXYZ. Координата Z - минимальная
        /// OOX
        /// OOO
        /// OOO
        /// </summary>
        private XYZ GetCurrentBBoxZMin_HRight(BoundingBoxXYZ boundingBox) => new XYZ(boundingBox.Max.X - 0.01, boundingBox.Max.Y - 0.01, boundingBox.Min.Z);

        /// <summary>
        /// Получить левую верхнюю точку BoundingBoxXYZ. Координата Z - минимальная
        /// XOO
        /// OOO
        /// OOO
        /// </summary>
        private XYZ GetCurrentBBoxZMin_HLeft(BoundingBoxXYZ boundingBox) => new XYZ(boundingBox.Max.X - 0.01, boundingBox.Min.Y + 0.01, boundingBox.Min.Z);

        /// <summary>
        /// Получить центральную точку BoundingBoxXYZ по плоскостям XY. Координата Z - минимальная
        /// OOO
        /// OXO
        /// OOO
        /// </summary>
        private XYZ GetCurrentBBoxZMin_Center(BoundingBoxXYZ boundingBox)
        {
            double centerX = (boundingBox.Min.X + boundingBox.Max.X) / 2.0;
            double centerY = (boundingBox.Min.Y + boundingBox.Max.Y) / 2.0;

            XYZ centerPoint = new XYZ(centerX, centerY, boundingBox.Min.Z);
            return centerPoint;
        }

        /// <summary>
        /// Получить правую нижнюю точку BoundingBoxXYZ. Координата Z - минимальная
        /// OOO
        /// OOO
        /// OOX
        /// </summary>
        private XYZ GetCurrentBBoxZMin_LRight(BoundingBoxXYZ boundingBox) => new XYZ(boundingBox.Min.X + 0.01, boundingBox.Max.Y - 0.01, boundingBox.Min.Z);

        /// <summary>
        /// Получить левую нижнюю точку BoundingBoxXYZ. Координата Z - минимальная
        /// OOO
        /// OOO
        /// XOO
        /// </summary>
        private XYZ GetCurrentBBoxZMin_LLeft(BoundingBoxXYZ boundingBox) => new XYZ(boundingBox.Min.X + 0.01, boundingBox.Min.Y + 0.01, boundingBox.Min.Z);

        /// <summary>
        /// Поделить BoundingBoxXYZ на части в плоскости XY
        /// </summary>
        /// <returns>Массив из N BoundingBoxXYZ</returns>
        private BoundingBoxXYZ[] SetCurrentBBoxArray(int count)
        {
            BoundingBoxXYZ[] bboxArray = new BoundingBoxXYZ[count];

            foreach (BoundingBoxXYZ bbox in MEPElemBBoxes)
            {
                // Calculate width and height of the bounding box in the XY plane
                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y;

                // Calculate the size of each piece
                double pieceWidth = width / count;
                double pieceHeight = height / count;

                for (int i = 0; i < count; i++)
                {
                    double minX = bbox.Min.X + (i * pieceWidth);
                    double minY = bbox.Min.Y + (i * pieceHeight);
                    double maxX = minX + pieceWidth;
                    double maxY = minY + pieceHeight;

                    bboxArray[i] = new BoundingBoxXYZ
                    {
                        Min = new XYZ(minX, minY, bbox.Min.Z),
                        Max = new XYZ(maxX, maxY, bbox.Max.Z)
                    };
                }
            }

            return bboxArray;
        }
    }
}
