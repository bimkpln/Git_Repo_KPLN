using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по помещениям АР и элементам для замера дистанции (перекрытия, лестницы)
    /// </summary>
    internal class CheckMEPHeightARData
    {
        public Solid _currentSolid;
        public BoundingBoxXYZ _currentBBox;
        /// <summary>
        /// Список части имен помещений, которые НЕ являются ошибками
        /// </summary>
        private static readonly List<string> _arRoomNameExceptionColl = new List<string>()
        {
            "итп",
            "пространство",
            "насосная",
            "камера"
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
        /// Solid помещения
        /// </summary>
        public Solid CurrentRoomSolid
        {
            get
            {
                if (_currentSolid == null)
                {
                    if (CurrentRoom == null) throw new Exception("Не определно помещение для анализа");
                    
                    GeometryElement roomShell = CurrentRoom.ClosedShell;
                    foreach (GeometryObject obj in roomShell)
                    {
                        Solid flSolid = obj as Solid;
                        if (flSolid != null && flSolid.Volume != 0)
                        {
                            _currentSolid = flSolid;
                            return _currentSolid;
                        }
                    }
                    throw new Exception($"Не удалось получить геометрию у помещения с id: {CurrentRoom.Id}"); ;
                }

                return _currentSolid;
            }
        }

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
                
                IEnumerable<Room> roomsColl = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>();
                
                foreach (Room room in roomsColl)
                {
                    // Игнорирую исключения
                    string rName = room.Name.ToLower();
                    if (_arRoomNameExceptionColl.Any(i => rName.Contains(i))) continue;

                    // Получаю АР - элементы согласно фильтрам и генерирую объекты
                    CheckMEPHeightARData arData = new CheckMEPHeightARData(room);
                
                    // Фильтр для поиска элементов, пересекающихся с Solid помещения
                    ElementIntersectsSolidFilter solidFilter = new ElementIntersectsSolidFilter(arData.CurrentRoomSolid);
                
                    // Фильтр для поиска элементов, входящих в BoundingBox помещения
                    BoundingBoxIntersectsFilter bboxFilter = arData.CreateFilter(-5);
                
                    // Объединяю фильтры
                    LogicalOrFilter finalFilter = new LogicalOrFilter(solidFilter, bboxFilter);

                    // Генерирую коллекцию элементов элементов в помещениях: Добавляю лестницы
                    arData.CurrentDownFaceElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Stairs)
                        .WhereElementIsNotElementType()
                        .WherePasses(finalFilter));

                    // Генерирую коллекцию элементов элементов в помещениях: Добавляю полы
                    arData.CurrentDownFaceElemsColl.AddRange(new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Floors)
                        .WhereElementIsNotElementType()
                        .WherePasses(finalFilter));

                    if (arData.CurrentDownFaceElemsColl.Count == 0)
                    {
                        Print($"Для помещения {arData.CurrentRoom.Name} - не удалось найти основания. Проверь элементы вручную", KPLN_Loader.Preferences.MessageType.Warning);
                        continue;
                    }

                    // Генерирую коллекцию поверхностей для анализа
                    foreach (Element elem in arData.CurrentDownFaceElemsColl)
                    {
                        FaceArray elemFaceArray = arData.GetElemsFaces(elem);
                        foreach (Face face in elemFaceArray)
                        {
                            arData.CurrentDownFacesArray.Append(face);
                        }
                    }

                    if (arData.CurrentDownFacesArray.IsEmpty)
                        throw new Exception($"Не удалось получить поверхности для анализа у помещения {arData.CurrentRoom.Id}");

                    result.Add(arData);
                }
            }

            return result;
        }

        /// <summary>
        /// Создание фильтра по BoundingBox
        /// </summary>
        /// <param name="Z_tolerance">Погрешность по оси Z </param>
        /// <returns></returns>
        private BoundingBoxIntersectsFilter CreateFilter(double Z_tolerance)
        {
            double minX = CurrentRoomBBox.Min.X;
            double minY = CurrentRoomBBox.Min.Y;

            double maxX = CurrentRoomBBox.Max.X;
            double maxY = CurrentRoomBBox.Max.Y;

            double sminX = Math.Min(minX, maxX);
            double sminY = Math.Min(minY, maxY);

            double smaxX = Math.Max(minX, maxX);
            double smaxY = Math.Max(minY, maxY);

            XYZ pntMax = new XYZ(smaxX, smaxY, CurrentRoomBBox.Max.Z);
            XYZ pntMin = new XYZ(sminX, sminY, CurrentRoomBBox.Min.Z + Z_tolerance);

            Outline outline = new Outline(pntMin, pntMax);

            return new BoundingBoxIntersectsFilter(outline);
        }

        /// <summary>
        /// Получить массив FaceArray
        /// </summary>
        private FaceArray GetElemsFaces(Element elem)
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
    }
}
