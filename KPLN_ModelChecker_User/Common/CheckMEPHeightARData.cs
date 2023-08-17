using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по помещениям АР и элементам для замера дистанции (перекрытия, лестницы)
    /// </summary>
    internal class CheckMEPHeightARData
    {
        public BoundingBoxXYZ _currentBBox;
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
        /// Коллекция элементов, которое включает в себя помещение
        /// </summary>
        public List<Element> CurrentDownFaceElemsColl { get; private set; } = new List<Element>();

        /// <summary>
        /// Коллекция поверхностей для проекции
        /// </summary>
        public FaceArray CurrentDownFacesArray { get; private set; } = new FaceArray();

        /// <summary>
        /// Генерация коллекции CheckMEPHeightARData
        /// </summary>
        /// <param name="linkInsts">Файлы АР для анализа</param>
        /// <param name="roomsColl">Набор элементов для селекции</param>
        /// <returns></returns>
        public static List<CheckMEPHeightARData> PreapareMEPHeightARDataColl(IEnumerable<RevitLinkInstance> linkInsts)
        {
            List<CheckMEPHeightARData> result = new List<CheckMEPHeightARData>();

            foreach (RevitLinkInstance rli in linkInsts)
            {
                Document linkDoc = rli.GetLinkDocument();

                // Коллекция помещений
                IEnumerable<Room> roomsColl = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>();

                // Коллекция элемнтов для проекции
                List<Element> downElemsColl = new List<Element>();

                // Добавляю лестницы
                downElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType()
                    .ToList());

                // Добавляю полы
                downElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .ToList());

                // Генерирую коллекцию поверхностей для анализа
                Dictionary<Element, FaceArray> arElemFaceArrayDict = new Dictionary<Element, FaceArray>(downElemsColl.Count());
                foreach (Element elem in downElemsColl)
                {
                    arElemFaceArrayDict.Add(elem, GetElemsFaces(elem));
                }

                foreach (Room room in roomsColl)
                {
                    // Игнорирую исключения
                    string rName = room.Name.ToLower();
                    if (_arRoomNameExceptionColl.Any(i => rName.Contains(i))) continue;

                    // Получаю АР - элементы согласно фильтрам и генерирую объекты
                    CheckMEPHeightARData arData = new CheckMEPHeightARData(room);

                    // Генерирую коллекцию поверхностей для анализа
                    foreach (KeyValuePair<Element, FaceArray> kvp in arElemFaceArrayDict)
                    {
                        if (arData.IsFloorInRoom(kvp.Key))
                        {
                            foreach (Face face in kvp.Value)
                            {
                                arData.CurrentDownFacesArray.Append(face);
                            }
                        }

                    }

                    if (arData.CurrentDownFacesArray.IsEmpty)
                        Print($"Для помещения {arData.CurrentRoom.Name} - не удалось найти основания. Проверь элементы вручную", MessageType.Warning);


                    //// Фильтр для поиска элементов, пересекающихся с BoundingBox помещения
                    //BoundingBoxIntersectsFilter bboxIntersectsFilter = new BoundingBoxIntersectsFilter(arData.CreateOutlineForFilter(-5));

                    //// Фильтр для поиска элементов, входящих в BoundingBox помещения
                    //BoundingBoxIsInsideFilter bboxIsInsideFilter = new BoundingBoxIsInsideFilter(arData.CreateOutlineForFilter(-5));

                    //// Объединяю фильтры
                    //LogicalOrFilter finalFilter = new LogicalOrFilter(bboxIntersectsFilter, bboxIsInsideFilter);

                    //// Генерирую коллекцию элементов элементов в помещениях: Добавляю лестницы
                    //arData.CurrentDownFaceElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                    //    .OfCategory(BuiltInCategory.OST_Stairs)
                    //    .WhereElementIsNotElementType()
                    //    .WherePasses(finalFilter));

                    //// Генерирую коллекцию элементов элементов в помещениях: Добавляю полы
                    //arData.CurrentDownFaceElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                    //    .OfCategory(BuiltInCategory.OST_Floors)
                    //    .WhereElementIsNotElementType()
                    //    .WherePasses(finalFilter));

                    //if (arData.CurrentDownFaceElemsColl.Count == 0)
                    //    Print($"Для помещения {arData.CurrentRoom.Name} - не удалось найти основания. Проверь элементы вручную", KPLN_Loader.Preferences.MessageType.Warning);


                    result.Add(arData);
                }
            }

            return result;
        }

        private bool IsFloorInRoom(Element floor)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };

            foreach (BoundarySegment segment in GetRoomBoundarySegments(CurrentRoom, options))
            {
                Curve curve = segment.GetCurve();
                XYZ pointOnFloor = curve.GetEndPoint(0);

                if (IsPointInsideFloor(floor, pointOnFloor))
                    return true;
            }

            return false;
        }

        private IEnumerable<BoundarySegment> GetRoomBoundarySegments(Room room, SpatialElementBoundaryOptions options)
        {
            IList<IList<BoundarySegment>> segments = room.GetBoundarySegments(options);
            List<BoundarySegment> allSegments = new List<BoundarySegment>();

            foreach (IList<BoundarySegment> boundarySegments in segments)
            {
                foreach (BoundarySegment segment in boundarySegments)
                {
                    allSegments.Add(segment);
                }
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
    }
}
