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
        private const int _bboxArrayCount = 3;
        private BoundingBoxXYZ _currentBBox;
        public Solid _currentSolid;
        private BoundingBoxXYZ[] _currentBBoxArray;
        /// <summary>
        /// Точность в проверке при анализе положения элементов
        /// </summary>
        private static readonly double _toleranceToCheck = 0.1;

        public CheckMEPHeightMEPData(Element elem)
        {
            CurrentElement = elem;
        }

        public Element CurrentElement { get; }

        public Solid CurrentSolid
        {
            get
            {
                if (_currentSolid == null)
                {
                    GeometryElement geomElem = CurrentElement.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });

                    foreach (GeometryObject gObj in geomElem)
                    {
                        if (gObj is Solid solid1)
                        {
                            _currentSolid = solid1;
                            break;
                        }
                        else if (gObj is GeometryInstance gInst)
                        {
                            GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                            double tempVolume = 0;
                            foreach (GeometryObject gObj2 in instGeomElem)
                            {
                                if (gObj2 is Solid solid2 && solid2.Volume > tempVolume)
                                    _currentSolid = solid2;
                                //else if (gObj2 is Solid solid3 && solid3.Volume > tempVolume)
                                //{
                                //    tempVolume = solid3.Volume;
                                //    _currentSolid = solid3;
                                //}
                            }
                        }
                    }

                    //if (_currentSolid == null)
                        //throw new Exception($"Не удалось получить геометрию у элемента с id: {CurrentElement.Id}");
                }

                return _currentSolid;
            }
        }

        public BoundingBoxXYZ CurrentBBox
        {
            get
            {
                if (_currentBBox == null)
                {

                    if (_currentSolid == null)
                        throw new Exception($"Не удалось получить геометрию у элемента с id: {CurrentElement.Id}");
                    
                    BoundingBoxXYZ bbox = CurrentSolid.GetBoundingBox();
                    Transform transform = bbox.Transform;
                    _currentBBox = new BoundingBoxXYZ()
                    {
                        Max = transform.OfPoint(bbox.Max),
                        Min = transform.OfPoint(bbox.Min),
                    };
                }

                return _currentBBox;
            }
        }

        /// <summary>
        /// Массив равнозначных в плоскости XY BoundingBoxXYZ (для проверки по всем типам помещений)
        /// </summary>
        public BoundingBoxXYZ[] CurrentBBoxArray
        {
            get
            {
                if (_currentBBoxArray == null)
                {
                    if (CurrentElement.Location is LocationCurve _) _currentBBoxArray = SetCurrentBBoxArray(_bboxArrayCount);
                    else _currentBBoxArray = new BoundingBoxXYZ[1] { CurrentBBox };
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
                                if (mepData.CurrentElement.Id.IntegerValue == elemToCheck.Id.IntegerValue)
                                    continue;

                                if (mepData.CurrentElement is FamilyInstance famInst)
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
        /// Проверка на нахождение элемента в указанном помещинии
        /// </summary>
        /// <param name="mepData">Спец. класс для проверки</param>
        /// <param name="arData">Спец. класс для проверки</param>
        public static bool IsElemInCurrentRoomCheck(CheckMEPHeightMEPData mepData, CheckMEPHeightARData arData)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };

            IList<BoundarySegment> segments = arData.CurrentRoom.GetBoundarySegments(options).SelectMany(bs => bs).ToList();
            
            foreach (BoundarySegment segment in segments)
            {
                Curve curve = segment.GetCurve();
                var solidIntResult = mepData.CurrentSolid.IntersectWithCurve(curve, new SolidCurveIntersectionOptions() { ResultType = SolidCurveIntersectionMode.CurveSegmentsInside});
                if (solidIntResult != null && solidIntResult.SegmentCount > 0)
                {
                    return true;
                }
            }

            return false;

            //if (mepData.CurrentSolid.Volume == 0)
            //    return false;
            
            //if (mepData.CurrentBBox.Max.X < arData.CurrentRoomBBox.Min.X || mepData.CurrentBBox.Min.X > arData.CurrentRoomBBox.Max.X)
            //    return false;

            //if (mepData.CurrentBBox.Max.Y < arData.CurrentRoomBBox.Min.Y || mepData.CurrentBBox.Min.Y > arData.CurrentRoomBBox.Max.Y)
            //    return false;

            //if (mepData.CurrentBBox.Max.Z < arData.CurrentRoomBBox.Min.Z || mepData.CurrentBBox.Min.Z > arData.CurrentRoomBBox.Max.Z)
            //    return false;

            //return true;
        }

        /// <summary>
        /// Проверить коллекцию эл-в ИОС нарушение высоты в помещении АР 
        /// </summary>
        /// <param name="currentRoomMEPDataColl">Коллекция спец. классов ИОС для проверки</param>
        /// <param name="arData">Спец. класс АР для проверки</param>
        public static CheckMEPHeightMEPData[] CheckIOSElemsForMinDistErrorByAR(CheckMEPHeightMEPData[] currentRoomMEPDataColl, CheckMEPHeightARData arData) =>
            currentRoomMEPDataColl.Where(m => m.IsHeigtError(arData)).ToArray();
        

        /// <summary>
        /// Проверить элемент на факт нарушения высоты
        /// </summary>
        /// <param name="arData">Спец. класс для проверки</param>
        private bool IsHeigtError(CheckMEPHeightARData arData)
        {
            // Проверка элементов на предмет пространственного положения выше 1.5 м болле чем на 1 часть
            if (CurrentBBoxArray.Where(b => b.Max.Z > 5).Count() > 0)
            {
                double tempIntDist = Double.MaxValue;
                foreach (BoundingBoxXYZ checkBBox in CurrentBBoxArray)
                {
                    #region Проверка вертикальных участков на вертикальную проходку между этажами (стояк). К ним отношу все участки, которые больше 1.5 м
                    if (CurrentElement.Location is LocationCurve locCurve)
                    {
                        if (locCurve.Curve is Line locLine)
                        {
                            XYZ direction = locLine.Direction;
                            if (Math.Abs(direction.X) < _toleranceToCheck && Math.Abs(direction.Y) < _toleranceToCheck && locCurve.Curve.ApproximateLength > 5)
                                return false;
                        }
                        else
                            throw new Exception($"Не удалось проанализировать Location элемента ИОС с id: {CurrentElement.Id}");
                    }
                    #endregion

                    #region Проверка на минимальную дистанцию до поверхности
                    //XYZ point = GetCurrentBBoxZMin_Center(checkBBox);
                    //foreach (Face face in arData.CurrentDownFacesArray)
                    //{
                    //    IntersectionResult intRes = face.Project(point);
                    //    if (intRes != null && intRes.Distance < tempIntDist)
                    //    {
                    //        tempIntDist = intRes.Distance;
                    //        double iosDistance = point.DistanceTo(intRes.XYZPoint);
                    //        if (iosDistance < arData.CurrentRoomMinDistance)
                    //            return true;
                    //    }
                    //}

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
                        //foreach (Face face in arData.CurrentDownFacesArray)
                        //{
                        //    IntersectionResult intRes = face.Project(point);
                        //    if (intRes != null && intRes.Distance < tempIntDist)
                        //    {
                        //        tempIntDist = intRes.Distance;
                        //        double iosDistance = point.DistanceTo(intRes.XYZPoint);
                        //        if (iosDistance < arData.CurrentRoomMinDistance)
                        //            return true;
                        //    }
                        //}
                    }
                    #endregion
                }
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

            // Calculate width and height of the bounding box in the XY plane
            double width = CurrentBBox.Max.X - CurrentBBox.Min.X;
            double height = CurrentBBox.Max.Y - CurrentBBox.Min.Y;

            // Calculate the size of each piece
            double pieceWidth = width / count;
            double pieceHeight = height / count;

            for (int i = 0; i < count; i++)
            {
                double minX = CurrentBBox.Min.X + (i * pieceWidth);
                double minY = CurrentBBox.Min.Y + (i * pieceHeight);
                double maxX = minX + pieceWidth;
                double maxY = minY + pieceHeight;

                bboxArray[i] = new BoundingBoxXYZ
                {
                    Min = new XYZ(minX, minY, CurrentBBox.Min.Z),
                    Max = new XYZ(maxX, maxY, CurrentBBox.Max.Z)
                };
            }

            return bboxArray;
        }
    }
}
