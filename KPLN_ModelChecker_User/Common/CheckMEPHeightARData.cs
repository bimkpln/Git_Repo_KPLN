using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по помещениям АР и элементам для замера дистанции (перекрытия, лестницы)
    /// </summary>
    internal class CheckMEPHeightARData
    {
        private BoundingBoxXYZ _currentBBox;
        private Solid _currentRoomSolid;
        /// <summary>
        /// Список части имен помещений, которые НЕ являются ошибками
        /// </summary>
        private static readonly List<string> _arRoomNameExceptionColl = new List<string>()
        {
            "итп",
            "пространство",
            "насосная",
            "камера",
        };
        /// <summary>
        /// Минимальная высота для проверки остальных помещений
        /// </summary>
        private const double _minRoomDistance = 6.562;
        /// <summary>
        /// Список вариантов названий лестничных клеток
        /// </summary>
        private static readonly List<string> _arRoomStairsNames = new List<string> { "лк", "лестничн", };
        /// <summary>
        /// Минимальная высота для проверки лестничных клеток
        /// </summary>
        private const double _minStairsDistance = 7.218;

        private BoundarySegment[] _currentRoomBoundSegmArr;

        private CheckMEPHeightARData(Room room)
        {
            CurrentRoom = room;
        }

        /// <summary>
        /// Помещение АР
        /// </summary>
        public Room CurrentRoom { get; private set; }

        /// <summary>
        /// BoundingBoxXYZ помещения
        /// </summary>
        public BoundingBoxXYZ CurrentRoomBBox
        {
            get
            {
                if (_currentBBox == null)
                {
                    if (CurrentRoom == null) throw new Exception("Не определно помещение для анализа");

                    _currentBBox = CurrentRoom.get_BoundingBox(null);
                    return _currentBBox;
                }

                return _currentBBox;
            }
        }

        public Solid CurrentRoomSolid
        {
            get
            {
                if (_currentRoomSolid == null)
                {
                    if (CurrentRoom == null) throw new Exception("Не определно помещение для анализа");

                    GeometryElement geomElem = CurrentRoom.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });
                    foreach (GeometryObject gObj in geomElem)
                    {
                        if (gObj is Solid solid)
                        {
                            _currentRoomSolid = solid;
                            break;
                        }
                        else
                            throw new Exception("Не определена геометрия помещения");
                    }
                }

                return _currentRoomSolid;
            }
        }

        /// <summary>
        /// Минимальная допустимая высота размещения элементов в данном помещении
        /// </summary>
        public double CurrentRoomMinDistance
        {
            get
            {
                if (CurrentRoom == null) throw new Exception("Не определно помещение для анализа");

                if (_arRoomStairsNames.Any(i => CurrentRoom.Name.ToLower().Contains(i)))
                    return _minStairsDistance;

                return _minRoomDistance;
            }
        }

        /// <summary>
        /// Коллекция элементов (нижние границы), которое включает в себя помещение
        /// </summary>
        public List<CheckMEPHeightARElemData> CurrentDownARElemDataColl { get; private set; } = new List<CheckMEPHeightARElemData>();

        /// <summary>
        /// Массив границ помещения
        /// </summary>
        private BoundarySegment[] CurrentRoomBoundSegmArr
        {
            get
            {
                if (_currentRoomBoundSegmArr == null)
                {
                    _currentRoomBoundSegmArr = CurrentRoom
                    .GetBoundarySegments(new SpatialElementBoundaryOptions { SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish })
                    .SelectMany(bs => bs)
                    .ToArray();
                }

                return _currentRoomBoundSegmArr;
            }
        }

        /// <summary>
        /// Генерация коллекции CheckMEPHeightARData
        /// </summary>
        /// <param name="linkInsts">Файлы АР для анализа</param>
        /// <returns></returns>
        public static List<CheckMEPHeightARData> PreapareMEPHeightARDataColl(IEnumerable<RevitLinkInstance> linkInsts)
        {
            List<CheckMEPHeightARData> result = new List<CheckMEPHeightARData>();

            // Коллекция помещений
            List<Room> roomsColl = new List<Room>();
            // Коллекция элемнтов для проекции
            List<CheckMEPHeightARElemData> projectionElemsColl = new List<CheckMEPHeightARElemData>();
            //Dictionary<Element, FaceArray> arElemFaceArrayDict = new Dictionary<Element, FaceArray>();

            // Анализ связей на помещения и геометрию
            for (int i = 0; i < linkInsts.Count(); i++)
            {
                RevitLinkInstance rli = linkInsts.ElementAt(i);
                Document linkDoc = rli.GetLinkDocument();

                // Коллекция помещений
                roomsColl.AddRange(new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToArray().Cast<Room>());

                // Коллекция элемнтов для проекции
                Element[] stairs = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToArray();
                Element[] floors = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToArray();
                projectionElemsColl
                    .AddRange(stairs
                        .Select(s => new CheckMEPHeightARElemData(s, rli)));
                projectionElemsColl
                    .AddRange(floors
                        .Where(f => f.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() > 0)
                        .Select(f => new CheckMEPHeightARElemData(f, rli)));
            }

            // Делю список помещений на равные части для ускорения
            int separCount = Environment.ProcessorCount / 3 * 2;
            //int separCount = 1;
            int chunkSize = (int)Math.Ceiling((double)roomsColl.Count / separCount);
            List<List<Room>> chunkRoomColl = roomsColl
                .Select((element, index) => new { Element = element, Index = index })
                .GroupBy(x => x.Index / chunkSize)
                .Select(group => group.Select(x => x.Element).ToList())
                .ToList();
            
            // Запускаю таски на разделенный список
            object lockerResult = new object();
            object lockerRoom = new object();
            Task[] arSolidTasks = new Task[separCount];
            for (int i = --separCount; i >= 0; i--)
            {
                foreach (List<Room> roomColl in chunkRoomColl)
                {
                    arSolidTasks[i] = Task.Factory.StartNew(() =>
                    {
                        foreach(Room room in roomColl)
                        {
                            CheckMEPHeightARData arData = new CheckMEPHeightARData(room);
                            lock (lockerRoom)
                            {
                                XYZ cntrPnt = arData.CurrentRoomSolid.ComputeCentroid();
                                foreach (CheckMEPHeightARElemData arElem in projectionElemsColl)
                                {
                                    if (arData.IsElemInCurrentRoom(arElem.ARElement.get_BoundingBox(null)))
                                    {
                                        if (arElem.ARElement is Floor floor)
                                        {
                                            XYZ upPrjPoint = floor.GetVerticalProjectionPoint(cntrPnt, FloorFace.Top);
                                            if (upPrjPoint != null && cntrPnt.Z > upPrjPoint.Z)
                                            {
                                                FaceArray faceArr = arElem.ARSolid.Faces;
                                                foreach (Face face in faceArr)
                                                {
                                                    IntersectionResult intRes = face.Project(cntrPnt);
                                                    if (intRes != null)
                                                    {
                                                        XYZ pntToCheck = new XYZ(intRes.XYZPoint.X, intRes.XYZPoint.Y, cntrPnt.Z);
                                                        if (arData.CurrentRoom.IsPointInRoom(pntToCheck))
                                                        {
                                                            arData.CurrentDownARElemDataColl.Add(arElem);
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            lock (lockerResult)
                            {
                                result.Add(arData);
                            }
                        }
                    });
                }
            }
            Task.WaitAll(arSolidTasks);

            return result;
        }

        private bool IsElemInCurrentRoom(BoundingBoxXYZ bbox) =>
            (CurrentRoomBBox.Max.X >= bbox.Min.X && CurrentRoomBBox.Min.X <= bbox.Max.X)
            && (CurrentRoomBBox.Max.Y >= bbox.Min.Y && CurrentRoomBBox.Min.Y <= bbox.Max.Y)
            && (CurrentRoomBBox.Max.Z >= bbox.Min.Z && CurrentRoomBBox.Min.Z <= bbox.Max.Z + 5);

        private static List<XYZ> RoomDownEdgeColl(CheckMEPHeightARData arData)
        {
            List<XYZ> roomDownEdgeColl = new List<XYZ>();
            
            LocationPoint roomLP = arData.CurrentRoom.Location as LocationPoint;
            XYZ roomPoint = roomLP.Point;
            double roomPoint_Z = roomPoint.Z;
            EdgeArray arDataEdgeArr = arData.CurrentRoomSolid.Edges;
            foreach (Edge edge in arDataEdgeArr)
            {
                XYZ[] tesselatedXYZ = edge.Tessellate().ToArray();
                if (tesselatedXYZ.All(t => Math.Abs(t.Z - roomPoint_Z) < 0.01))
                    roomDownEdgeColl.AddRange(tesselatedXYZ);
            }

            return roomDownEdgeColl;
        }

        private bool IsElemInRoom(Element elem)
        {
            BoundingBoxXYZ elemBbox = elem.get_BoundingBox(null);
            if (elemBbox != null)
            {
                if ((CurrentRoomBBox.Min.X <= elemBbox.Min.X && CurrentRoomBBox.Min.Y <= elemBbox.Min.Y && CurrentRoomBBox.Min.Z - 5 <= elemBbox.Min.X)
                    && (CurrentRoomBBox.Max.X >= elemBbox.Max.X && CurrentRoomBBox.Max.Y >= elemBbox.Max.Y && CurrentRoomBBox.Max.Z >= elemBbox.Max.X + 5))
                    return false;

                foreach (BoundarySegment segment in CurrentRoomBoundSegmArr)
                {
                    Curve curve = segment.GetCurve();
                    XYZ pointOnBoundSegm = curve.GetEndPoint(0);
                    if (pointOnBoundSegm.X >= elemBbox.Min.X && pointOnBoundSegm.X <= elemBbox.Max.X &&
                        pointOnBoundSegm.Y >= elemBbox.Min.Y && pointOnBoundSegm.Y <= elemBbox.Max.Y &&
                        pointOnBoundSegm.Z >= elemBbox.Min.Z - 5 && pointOnBoundSegm.Z - 5 <= elemBbox.Max.Z)
                    {
                        return true;
                    }
                }
            }


            //SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions
            //{
            //    SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            //};

            //IList<BoundarySegment> segments = CurrentRoom.GetBoundarySegments(options).SelectMany(bs => bs).ToList();
            //foreach (BoundarySegment segment in segments)
            //{
            //    Curve curve = segment.GetCurve();
            //    XYZ pointOnBoundSegm = curve.GetEndPoint(0);

            //    BoundingBoxXYZ elemBbox = floor.get_BoundingBox(null);
            //    if (elemBbox != null)
            //    {
            //        XYZ min = elemBbox.Min;
            //        XYZ max = elemBbox.Max;

            //        if (pointOnBoundSegm.X >= min.X && pointOnBoundSegm.X <= max.X &&
            //            pointOnBoundSegm.Y >= min.Y && pointOnBoundSegm.Y <= max.Y &&
            //            pointOnBoundSegm.Z >= min.Z - 5 && pointOnBoundSegm.Z - 5 <= max.Z)
            //        {
            //            return true;
            //        }
            //    }
            //}

            //BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
            //if (bbox != null)
            //{
            //    if (bbox.Max.X < CurrentRoomBBox.Min.X || bbox.Min.X > CurrentRoomBBox.Max.X)
            //        return false;

            //    if (bbox.Max.Y < CurrentRoomBBox.Min.Y || bbox.Min.Y > CurrentRoomBBox.Max.Y)
            //        return false;

            //    if (bbox.Max.Z < CurrentRoomBBox.Min.Z || bbox.Min.Z > CurrentRoomBBox.Max.Z)
            //        return false;

            //    return true;
            //}

            return false;
        }

        private IEnumerable<BoundarySegment> GetRoomBoundarySegments(Room room, SpatialElementBoundaryOptions options)
        {
            IList<IList<BoundarySegment>> segments = room.GetBoundarySegments(options);
            List<BoundarySegment> allSegments = new List<BoundarySegment>();

            foreach (IList<BoundarySegment> boundarySegments in segments)
            {
                allSegments.AddRange(boundarySegments);
            }

            return allSegments;
        }

        private bool IsPointInsideFloor(Element floor, XYZ point)
        {
            BoundingBoxXYZ bbox = floor.get_BoundingBox(null);
            if (bbox != null)
            {
                XYZ min = bbox.Min;
                XYZ max = bbox.Max;

                if (point.X >= min.X && point.X <= max.X &&
                    point.Y >= min.Y && point.Y <= max.Y &&
                    point.Z >= min.Z - 5 && point.Z - 5 <= max.Z)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Получить массив FaceArray
        /// </summary>
        private static FaceArray GetElemsFaces(Element elem)
        {
            FaceArray result = new FaceArray();

            IList<Solid> elemSolids = new List<Solid>();
            GeometryElement geometryElement = elem.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });
            if (geometryElement != null)
            {
                GetSolidsFromGeomElem(geometryElement, Transform.Identity, elemSolids);
            }

            foreach (Solid solid in elemSolids)
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

                    result.Append(face);
                }
            }

            return result;
        }

        private static void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solids)
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

        /// <summary>
        /// Создание Outline для фиьлтров типа ElementQuickFilter
        /// </summary>
        /// <param name="Z_tolerance">Погрешность по оси Z </param>
        /// <returns></returns>
        private Outline CreateOutlineForFilter(Transform transform, double Z_tolerance)
        {
            double minX = CurrentRoomBBox.Min.X;
            double minY = CurrentRoomBBox.Min.Y;

            double maxX = CurrentRoomBBox.Max.X;
            double maxY = CurrentRoomBBox.Max.Y;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);

            Outline outline = new Outline(
                transform.Inverse.OfPoint(new XYZ(sminX, sminY, CurrentRoomBBox.Min.Z + Z_tolerance)),
                transform.Inverse.OfPoint(new XYZ(smaxX , smaxY, CurrentRoomBBox.Max.Z)));

            return outline;
        }
    }
}
