using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckMEPHeight
{
    /// <summary>
    /// Контейнер для сбора информации по помещениям АР и элементам для замера дистанции (перекрытия, лестницы)
    /// </summary>
    public class CheckMEPHeightARRoomData
    {
        /// <summary>
        /// От данного парамтера зависит точность опредления привязки элемента к помещению
        /// </summary>
        private const int _bboxExpanded = 10;
        private Transform _roomLinkTrans;
        private BoundingBoxXYZ _roomBBox;
        private Solid _roomSolid;

        /// <summary>
        /// Инициализация дефолтного класса
        /// </summary>
        private CheckMEPHeightARRoomData(Room room, RevitLinkInstance roomLinkInst, double currentRoomMinElemElevationForCheck, double currentRoomMinDistance)
        {
            CurrentRoom = room;
            RoomLinkInst = roomLinkInst;
            CurrentRoomMinElemElevationForCheck = currentRoomMinElemElevationForCheck;
            CurrentRoomMinDistance = currentRoomMinDistance;
        }

        /// <summary>
        /// Помещение АР
        /// </summary>
        public Room CurrentRoom { get; private set; }

        public RevitLinkInstance RoomLinkInst { get; set; }

        /// <summary>
        /// Минимальная допустимая высота размещения элементов в данном помещении
        /// </summary>
        public double CurrentRoomMinDistance { get; private set; }

        /// <summary>
        /// Мин отметка, при которой элемент считается с потенциальным нарушением
        /// </summary>
        public double CurrentRoomMinElemElevationForCheck { get; private set; }

        public Transform RoomLinkTrans
        {
            get
            {
                if (_roomLinkTrans == null)
                    _roomLinkTrans = RoomLinkInst.GetTotalTransform();

                return _roomLinkTrans;
            }
        }

        /// <summary>
        /// BoundingBoxXYZ помещения
        /// </summary>
        public BoundingBoxXYZ RoomBBox
        {
            get
            {
                if (_roomBBox == null)
                {
                    if (CurrentRoom == null) throw new Exception("Не определно помещение для анализа");

                    Options opt = new Options
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                        ComputeReferences = true
                    };
                    GeometryElement geometryElement = CurrentRoom.get_Geometry(opt);
                    foreach (GeometryObject geomObject in geometryElement)
                    {
                        switch (geomObject)
                        {
                            case Solid solid:
                                _roomSolid = solid;
                                return _roomBBox = GetBoundingBoxXYZ(solid);
                            case GeometryInstance geomInstance:
                                GeometryElement instGeomElem = geomInstance.GetInstanceGeometry();
                                return _roomBBox = GetBoundingBoxXYZ(instGeomElem);
                            case GeometryElement geomElem:
                                return _roomBBox = GetBoundingBoxXYZ(geomElem);
                        }
                    }
                }

                return _roomBBox;
            }
        }

        public Solid RoomSolid
        {
            get
            {
                if (_roomSolid == null)
                {
                    if (CurrentRoom == null)
                        throw new Exception("Не определно помещение для анализа");

                    Options opt = new Options
                    {
                        DetailLevel = ViewDetailLevel.Fine,
                        ComputeReferences = true
                    };
                    GeometryElement geomElem = CurrentRoom.get_Geometry(opt);
                    foreach (GeometryObject gObj in geomElem)
                    {
                        if (gObj is Solid solid)
                        {
                            _roomSolid = solid;
                            break;
                        }
                        else
                            throw new Exception("Не определена геометрия помещения");
                    }
                }

                return _roomSolid;
            }
        }

        /// <summary>
        /// Коллекция элементов (нижние границы), которое включает в себя помещение
        /// </summary>
        public List<CheckMEPHeightARElemData> RoomDownARElemDataColl { get; private set; } = new List<CheckMEPHeightARElemData>();

        /// <summary>
        /// Генерация коллекции CheckMEPHeightViewModel для стартового окна
        /// </summary>
        /// <param name="linkInsts">Файлы АР для анализа</param>
        public static CMHEntity[] PreapareMEPHeightViewModelDataDataColl(RevitLinkInstance[] linkInsts)
        {
            List<CMHEntity> result = new List<CMHEntity>();

            // Анализ связей на помещения
            foreach (RevitLinkInstance rli in linkInsts)
            {
                Document linkDoc = rli.GetLinkDocument();
                result.AddRange(new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Select(r => new CMHEntity(r)));
            }

            // Группирую по имени помещения и назначению помещения, чтобы были уникальные комбинации
            return result
                .GroupBy(r => new { r.VMCurrentRoomName, r.VMCurrentRoomDepartmentName })
                .Select(g => g.FirstOrDefault())
                .OrderBy(v => v.VMCurrentRoomName)
                .ToArray();
        }


        /// <summary>
        /// Генерация коллекции CheckMEPHeightARData по данным CheckMEPHeightViewModel
        /// </summary>
        /// <param name="vModels">Коллекция CheckMEPHeightViewModel для фильтрации</param>
        public static CheckMEPHeightARRoomData[] PreapareMEPHeightARRoomDataColl(IEnumerable<CMHEntity> vModels, RevitLinkInstance[] linkInsts)
        {
            List<CheckMEPHeightARRoomData> result = new List<CheckMEPHeightARRoomData>();

            // Коллекция для обработки
            HashSet<CheckMEPHeightARElemData> projectionElemsColl = new HashSet<CheckMEPHeightARElemData>();

            // Анализ связей на помещения и геометрию
            foreach (CMHEntity vm in vModels)
            {
                foreach (RevitLinkInstance rli in linkInsts)
                {
                    Document linkDoc = rli.GetLinkDocument();

                    // Помещения их линка
                    Room[] docRooms = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType()
                        .Cast<Room>()
                        .ToArray();

                    // Коллекция помещений с фильтрацией от ViewModel
                    foreach(Room r in docRooms)
                    {
                        string rName = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                        string rDeaprtment = r.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString() ?? string.Empty;
                        if (rName == vm.VMCurrentRoomName && rDeaprtment == vm.VMCurrentRoomDepartmentName)
                            result.Add(new CheckMEPHeightARRoomData(r, rli, vm.VMCurrentRoomMinElemElevationForCheck, vm.VMCurrentRoomMinDistance));
                    }

                    // Коллекция элементов для проекции (лестницы)
                    Element[] stairs = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Stairs).WhereElementIsNotElementType().ToArray();
                    foreach (Element stair in stairs)
                    {
                        projectionElemsColl.Add(new CheckMEPHeightARElemData(stair, rli));
                    }

                    // Коллекция элементов для проекции (перекрытия)
                    Element[] floors = new FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToArray();
                    foreach (Element floor in floors)
                    {
                        if (floor.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() > 0)
                            projectionElemsColl.Add(new CheckMEPHeightARElemData(floor, rli));
                    }
                }
            }

            // Подготовка элементов для работы
            for (int i = 0; i < result.Count; i++)
            {
                CheckMEPHeightARRoomData arRoomData = result.ElementAt(i);
                CheckMEPHeightARElemData[] quickFilteredArr = projectionElemsColl.Where(e => arRoomData.IsElemInCurrentRoom_QuickFilter(e.ARElemBBoxes)).ToArray();
                foreach (CheckMEPHeightARElemData arElem in quickFilteredArr)
                {
                    // Собираю все граничные и поглащенные элементы
                    arRoomData.SetRoomDownARElemsDataColl_ByIncludedToRoom(arElem);
                    arRoomData.SetRoomDownARElemsDataColl_ByProjectedWithRoom(arElem);
                }

                if (arRoomData.RoomDownARElemDataColl.Count == 0)
                    throw new CheckerException($"Не удалось получить нижние границы помещения с id: {arRoomData.CurrentRoom.Id} в связи: {arRoomData.RoomLinkInst.Name}");
            }

            return result.ToArray();
        }

        private BoundingBoxXYZ GetBoundingBoxXYZ(GeometryElement geomElem)
        {
            foreach (GeometryObject obj in geomElem)
            {
                Solid solid = obj as Solid;
                return GetBoundingBoxXYZ(solid);
            }

            return null;
        }

        private BoundingBoxXYZ GetBoundingBoxXYZ(Solid solid)
        {
            if (solid != null && solid.Volume != 0)
            {
                BoundingBoxXYZ bbox = solid.GetBoundingBox();
                Transform transform = bbox.Transform;
                Transform resultTrans = transform * RoomLinkTrans;
                return new BoundingBoxXYZ()
                {
                    Max = resultTrans.OfPoint(bbox.Max),
                    Min = resultTrans.OfPoint(bbox.Min),
                };
            }

            return null;
        }

        /// <summary>
        /// Внесение элементов в коллекцию RoomDownARElemDataColl, которые полностью погружены в текущее помещение, или полностью под ним
        /// </summary>
        /// <param name="arElem">Элемент АР для проверки</param>
        private void SetRoomDownARElemsDataColl_ByIncludedToRoom(CheckMEPHeightARElemData arElem)
        {
            foreach (Solid arElemSolid in arElem.ARElemSolids)
            {
                BoundingBoxXYZ arElemSolidBbox = arElemSolid.GetBoundingBox();
                Transform arElemSolidTransform = arElemSolidBbox.Transform;
                XYZ arElemNativeCntrPnt = arElemSolidTransform.OfPoint((arElemSolidBbox.Max + arElemSolidBbox.Min) / 2);

                // Анализ рабочих точек
                XYZ arElemCntrPnt = null;
                // Проверка Transform между помещением и геометрией (могут быть из разных линков, тогда Transform нужен, иначе - нет)
                if (arElem.ARElemLinkTrans.AlmostEqual(RoomLinkTrans))
                    arElemCntrPnt = arElemNativeCntrPnt;
                else
                    arElemCntrPnt = arElem.ARElemLinkTrans.OfPoint(arElemNativeCntrPnt);

                XYZ arElemCntrPntUpper_z1 = new XYZ(arElemCntrPnt.X, arElemCntrPnt.Y, arElemCntrPnt.Z + _bboxExpanded / 5);
                XYZ arElemCntrPntUpper_z2 = new XYZ(arElemCntrPnt.X, arElemCntrPnt.Y, arElemCntrPnt.Z + _bboxExpanded / 2);
                if (CurrentRoom.IsPointInRoom(arElemCntrPnt)
                    || CurrentRoom.IsPointInRoom(arElemCntrPntUpper_z1)
                    || CurrentRoom.IsPointInRoom(arElemCntrPntUpper_z2))
                {
                    RoomDownARElemDataColl.Add(arElem);
                    return;
                }
            }
        }

        /// <summary>
        /// Внесение элементов в коллекцию RoomDownARElemDataColl, которые выходят за грани текущего помещения, но проецируются в помещение
        /// </summary>
        /// <param name="arElem">Элемент АР для проверки</param>
        private void SetRoomDownARElemsDataColl_ByProjectedWithRoom(CheckMEPHeightARElemData arElem)
        {
            XYZ cntrPnt = RoomSolid.ComputeCentroid();
            FaceArray roomFaceArray = CheckMEPHeightARElemData.GetHorizontalDownFacesFromArray(RoomSolid.Faces, false);
            foreach (Face elemFace in arElem.ARElemDownFacesArray)
            {
                Mesh elemMesh = elemFace.Triangulate(1);
                if (elemMesh != null && elemMesh.Vertices.Any(v => v.Z <= cntrPnt.Z))
                {
                    foreach (Face roomFace in roomFaceArray)
                    {
                        Mesh roomMesh = roomFace.Triangulate(1);
                        IList<XYZ> roomVerts = roomMesh.Vertices;
                        var nativeBbox_Max = RoomLinkTrans.Inverse.OfPoint(RoomBBox.Max);
                        var nativeBbox_Min = RoomLinkTrans.Inverse.OfPoint(RoomBBox.Min);

                        foreach (XYZ roomVert in roomVerts)
                        {
                            // Для пограничных в помещениямх элементов - необходимо слегка смещать координату проверки
                            XYZ zipRoomVert = null;
                            if (roomVert.IsAlmostEqualTo(nativeBbox_Max, 0.1))
                                zipRoomVert = new XYZ(
                                    roomVert.X >= 0 ? roomVert.X - 0.25 : roomVert.X + 0.25,
                                    roomVert.Y >= 0 ? roomVert.Y - 0.25 : roomVert.Y + 0.25,
                                    roomVert.Z);
                            else if (roomVert.IsAlmostEqualTo(nativeBbox_Min, 0.1))
                                zipRoomVert = new XYZ(
                                    roomVert.X >= 0 ? roomVert.X + 0.25 : roomVert.X - 0.25,
                                    roomVert.Y >= 0 ? roomVert.Y + 0.25 : roomVert.Y - 0.25,
                                    roomVert.Z);
                            else
                                zipRoomVert = roomVert;

                            IntersectionResult intRes = elemFace.Project(zipRoomVert);
                            if (intRes != null && intRes.Distance <= CurrentRoom.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble())
                            {
                                // Анализ рабочих точек
                                XYZ pntToCheck = new XYZ(intRes.XYZPoint.X, intRes.XYZPoint.Y, cntrPnt.Z);
                                if (CurrentRoom.IsPointInRoom(pntToCheck))
                                {
                                    RoomDownARElemDataColl.Add(arElem);
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Определение находиться ли элемент (BoundingBoxXYZ) в границах (BoundingBoxXYZ) помещения
        /// </summary>
        /// <param name="bboxColl">BoundingBoxXYZ элементы для проверки</param>
        private bool IsElemInCurrentRoom_QuickFilter(List<BoundingBoxXYZ> bboxColl)
        {
            foreach (BoundingBoxXYZ bbox in bboxColl)
            {
                if ((RoomBBox.Max.X >= bbox.Min.X && RoomBBox.Min.X <= bbox.Max.X)
                    && (RoomBBox.Max.Y >= bbox.Min.Y && RoomBBox.Min.Y <= bbox.Max.Y)
                    && (RoomBBox.Max.Z >= bbox.Min.Z && RoomBBox.Min.Z <= bbox.Max.Z + _bboxExpanded))
                    return true;
            }

            return false;
        }
    }
}
