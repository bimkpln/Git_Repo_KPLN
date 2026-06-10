using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ApartmentManager.Common;
using KPLN_ApartmentManager.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private int PlaceRoomGeometryInTransaction(Document doc, List<PreparedApartmentWalls> preparedApartments,
            List<PreparedApartmentRooms> preparedRoomsByApartment,
            Level roomLevel, List<RoomAreaMismatchInfo> roomAreaMismatches,
            Dictionary<long, ApartmentProcessState> apartmentStates, ViewPlan roomPlan, List<string> debugMessages)
        {
            if (doc == null || preparedApartments == null || preparedApartments.Count == 0 || roomLevel == null)
                return 0;

            int createdRoomsCount = 0;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Геометрия помещений"))
            {
                t.Start();
                ApplyApartmentFailureHandling(t);

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null)
                        continue;

                    PreparedApartmentRooms apartmentRooms = preparedRoomsByApartment != null
                        ? preparedRoomsByApartment.FirstOrDefault(x => x != null && x.ApartmentId == apartmentWalls.ApartmentId)
                        : null;

                    ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);

                    int createdRoomsForApartment = 0;
                    int roomsForApartmentCount = apartmentRooms != null && apartmentRooms.Rooms != null ? apartmentRooms.Rooms.Count : 0;

                    int mismatchesBefore = roomAreaMismatches != null ? roomAreaMismatches.Count : 0;
                    List<ElementId> createdRoomIds = new List<ElementId>();

                    int createdRoomSeparators = PlaceRoomSeparatorsForApartment(
                        doc,
                        apartmentWalls,
                        apartmentWalls.RoomSeparatorLines,
                        roomLevel,
                        roomPlan,
                        state,
                        debugMessages);
                    if (createdRoomSeparators > 0)
                    {
                        state.CreatedRoomSeparatorsCount += createdRoomSeparators;
                        doc.Regenerate();
                    }

                    if (apartmentRooms != null && apartmentRooms.Rooms != null && apartmentRooms.Rooms.Count > 0)
                    {
                        createdRoomsForApartment = PlaceRoomsForApartment(
                            doc,
                            apartmentRooms,
                            roomLevel,
                            roomAreaMismatches,
                            createdRoomIds);
                    }

                    foreach (ElementId createdRoomId in createdRoomIds)
                        AddCreatedElementCandidate(state, createdRoomId);

                    createdRoomsCount += createdRoomsForApartment;

                    if (createdRoomsForApartment > 0)
                        state.HasCreatedRooms = true;

                    int skippedRoomsForApartment = roomsForApartmentCount - createdRoomsForApartment;
                    if (skippedRoomsForApartment > 0)
                        state.SkippedRoomsCount += skippedRoomsForApartment;

                    if (roomAreaMismatches != null && roomAreaMismatches.Count > mismatchesBefore)
                        state.HasRoomAreaMismatch = true;
                }

                doc.Regenerate();
                t.Commit();
            }

            return createdRoomsCount;
        }

        private int PlaceRoomSeparatorsForApartment(Document doc, PreparedApartmentWalls apartmentWalls, List<Line> roomSeparatorLines,
            Level roomLevel, ViewPlan roomPlan, ApartmentProcessState state, List<string> debugMessages)
        {
            if (doc == null || apartmentWalls == null || roomSeparatorLines == null ||
                roomSeparatorLines.Count == 0 || roomLevel == null || roomPlan == null)
            {
                return 0;
            }

            SketchPlane sketchPlane = null;
            try
            {
                Plane plane = Plane.CreateByNormalAndOrigin(
                    XYZ.BasisZ,
                    new XYZ(0, 0, roomLevel.Elevation));
                sketchPlane = SketchPlane.Create(doc, plane);
            }
            catch (Exception ex)
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Для квартиры ID = " + IDHelper.ElIdValue(apartmentWalls.ApartmentId) +
                    " не удалось создать рабочую плоскость разделителей помещений: " + ex.Message);
                return 0;
            }

            int createdCount = 0;
            foreach (Line sourceLine in roomSeparatorLines)
            {
                if (sourceLine == null || sourceLine.Length < IDHelper.ConvertMmToInternal(10))
                    continue;

                try
                {
                    XYZ p0 = sourceLine.GetEndPoint(0);
                    XYZ p1 = sourceLine.GetEndPoint(1);
                    Line projectedLine = Line.CreateBound(
                        new XYZ(p0.X, p0.Y, roomLevel.Elevation),
                        new XYZ(p1.X, p1.Y, roomLevel.Elevation));

                    CurveArray curves = new CurveArray();
                    curves.Append(projectedLine);

                    ModelCurveArray createdCurves = doc.Create.NewRoomBoundaryLines(sketchPlane, curves, roomPlan);
                    if (createdCurves == null)
                        continue;

                    foreach (ModelCurve modelCurve in createdCurves)
                    {
                        if (modelCurve == null)
                            continue;

                        createdCount++;
                        AddCreatedElementCandidate(state, modelCurve.Id);
                    }
                }
                catch (Exception ex)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Для квартиры ID = " + IDHelper.ElIdValue(apartmentWalls.ApartmentId) +
                        " не удалось создать разделитель помещения: " + ex.Message);
                }
            }

            return createdCount;
        }

        private PreparedApartmentRooms PrepareRoomsForApartment(Document doc, FamilyInstance apartmentFi, List<Line> loggiaAxisLines,
            List<Line> apartmentAxisLines, List<Line> shaftAxisLines, List<Line> roomSeparatorLines, View geometryView, List<string> debugMessages)
        {
            PreparedApartmentRooms result = new PreparedApartmentRooms();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;
            result.HasRoomSeparators = roomSeparatorLines != null && roomSeparatorLines.Count > 0;
            result.HasShafts = shaftAxisLines != null && shaftAxisLines.Count > 0;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
            foreach (FamilyInstance roomFi in roomInstances ?? new List<FamilyInstance>())
            {
                if (roomFi == null)
                    continue;

                try
                {
                    string roomName = GetRoomCategoryLabel(roomFi);
                    if (string.IsNullOrWhiteSpace(roomName))
                        roomName = "Помещение";

                    double expectedAreaInternal = 0;
                    TryGetAreaParamFromElementOrType(roomFi, out expectedAreaInternal, "КП_Р_Площадь", "КП_Р_ПЛОЩАДЬ");

                    RoomBoundaryLoopData loopData;
                    string diagnostic;
                    if (!TryBuildRoomBoundaryLoopFromInstanceGeometry(roomFi, geometryView, out loopData, out diagnostic) ||
                        loopData == null ||
                        loopData.Vertices == null ||
                        loopData.Vertices.Count < 3)
                    {
                        if (debugMessages != null)
                        {
                            debugMessages.Add(
                                "Для вложенного помещения ID = " +
                                IDHelper.ElIdValue(roomFi.Id) +
                                " не удалось собрать контур для точки помещения." +
                                (string.IsNullOrWhiteSpace(diagnostic) ? "" : " Диагностика: " + diagnostic + "."));
                        }

                        continue;
                    }

                    XYZ insertPointInProject = FindInteriorPointForRoomLoop(loopData.Vertices);
                    if (insertPointInProject == null)
                    {
                        if (debugMessages != null)
                        {
                            debugMessages.Add(
                                "Для вложенного помещения ID = " +
                                IDHelper.ElIdValue(roomFi.Id) +
                                " не удалось определить внутреннюю точку из геометрии экземпляра.");
                        }

                        continue;
                    }

                    result.Rooms.Add(new PreparedRoomPlacement
                    {
                        ApartmentId = apartmentFi.Id,
                        RoomName = roomName.Trim(),
                        InsertPoint = insertPointInProject,
                        ExpectedAreaInternal = expectedAreaInternal,
                        HasShaftInside = RoomContainsAnyLine(loopData.Vertices, shaftAxisLines)
                    });
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add(
                            "Ошибка подготовки помещения ID = " +
                            IDHelper.ElIdValue(roomFi.Id) + ": " + ex.Message);
                }
            }

            AddSyntheticLoggiaRoomPlacements(result, apartmentFi, doc, loggiaAxisLines, apartmentAxisLines);

            return result;
        }

        private static bool RoomContainsAnyLine(List<XYZ> roomVertices, List<Line> lines)
        {
            if (roomVertices == null || roomVertices.Count < 3 || lines == null || lines.Count == 0)
                return false;

            double tolerance = IDHelper.ConvertMmToInternal(20);

            foreach (Line line in lines)
            {
                if (line == null || line.Length < IDHelper.ConvertMmToInternal(10))
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ mid = new XYZ(
                    0.5 * (p0.X + p1.X),
                    0.5 * (p0.Y + p1.Y),
                    0.5 * (p0.Z + p1.Z));

                if (IsPointInsidePolygon2D(p0, roomVertices, tolerance) ||
                    IsPointInsidePolygon2D(p1, roomVertices, tolerance) ||
                    IsPointInsidePolygon2D(mid, roomVertices, tolerance))
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddSyntheticLoggiaRoomPlacements(PreparedApartmentRooms result, FamilyInstance apartmentFi, Document doc,
            List<Line> loggiaAxisLines, List<Line> apartmentAxisLines)
        {
            if (result == null || apartmentFi == null)
                return;

            if (result.Rooms != null && result.Rooms.Any(x => x != null && IsLoggiaRoomName(x.RoomName)))
                return;

            if (result.Rooms == null)
                result.Rooms = new List<PreparedRoomPlacement>();

            if (loggiaAxisLines == null || loggiaAxisLines.Count == 0)
                return;

            XYZ apartmentInteriorPoint = GetApartmentInteriorReferencePoint(apartmentFi, doc);
            double roomPointOffset = IDHelper.ConvertMmToInternal(300);
            double duplicateTol = IDHelper.ConvertMmToInternal(100);

            foreach (Line loggiaAxis in loggiaAxisLines)
            {
                if (loggiaAxis == null || loggiaAxis.Length < IDHelper.ConvertMmToInternal(10))
                    continue;

                XYZ p0 = loggiaAxis.GetEndPoint(0);
                XYZ p1 = loggiaAxis.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                XYZ normal = new XYZ(-dir.Y, dir.X, 0);
                XYZ mid = new XYZ(
                    0.5 * (p0.X + p1.X),
                    0.5 * (p0.Y + p1.Y),
                    0.5 * (p0.Z + p1.Z));

                XYZ roomPoint = BuildLoggiaRoomPointBetweenWalls(loggiaAxis, apartmentAxisLines);

                XYZ loggiaDir = normal;
                if (apartmentInteriorPoint != null)
                {
                    XYZ toInterior = Normalize2D(apartmentInteriorPoint - mid);
                    if (toInterior != null && Dot2D(loggiaDir, toInterior) > 0)
                        loggiaDir = new XYZ(-loggiaDir.X, -loggiaDir.Y, 0);
                }

                if (roomPoint == null)
                {
                    roomPoint = new XYZ(
                        mid.X + loggiaDir.X * roomPointOffset,
                        mid.Y + loggiaDir.Y * roomPointOffset,
                        mid.Z);
                }

                bool duplicate = result.Rooms.Any(x =>
                    x != null &&
                    x.InsertPoint != null &&
                    Distance2D(x.InsertPoint, roomPoint) <= duplicateTol);

                if (duplicate)
                    continue;

                result.Rooms.Add(new PreparedRoomPlacement
                {
                    ApartmentId = result.ApartmentId,
                    RoomName = "Лоджия",
                    InsertPoint = roomPoint,
                    ExpectedAreaInternal = 0
                });
            }
        }

        private static XYZ BuildLoggiaRoomPointBetweenWalls(Line loggiaAxis, List<Line> apartmentAxisLines)
        {
            if (loggiaAxis == null || apartmentAxisLines == null || apartmentAxisLines.Count == 0)
                return null;

            XYZ p0 = loggiaAxis.GetEndPoint(0);
            XYZ p1 = loggiaAxis.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);
            if (dir == null)
                return null;

            XYZ mid = new XYZ(
                0.5 * (p0.X + p1.X),
                0.5 * (p0.Y + p1.Y),
                0.5 * (p0.Z + p1.Z));

            XYZ normal = new XYZ(-dir.Y, dir.X, 0);
            double minProjectionOverlap = Math.Min(loggiaAxis.Length * 0.2, IDHelper.ConvertMmToInternal(300));
            double bestDistance = double.MaxValue;
            XYZ bestProjectedPoint = null;

            foreach (Line candidate in apartmentAxisLines)
            {
                if (candidate == null || candidate.Length < IDHelper.ConvertMmToInternal(10))
                    continue;

                XYZ c0 = candidate.GetEndPoint(0);
                XYZ c1 = candidate.GetEndPoint(1);
                XYZ candidateDir = Normalize2D(c1 - c0);
                if (candidateDir == null)
                    continue;

                if (Math.Abs(Dot2D(dir, candidateDir)) < 0.9)
                    continue;

                double signedDistance = Dot2D(c0 - mid, normal);
                double distance = Math.Abs(signedDistance);
                if (distance < IDHelper.ConvertMmToInternal(300))
                    continue;

                if (!HasParallelProjectionOverlap(loggiaAxis, candidate, minProjectionOverlap))
                    continue;

                XYZ projectedPoint = new XYZ(
                    mid.X + normal.X * signedDistance,
                    mid.Y + normal.Y * signedDistance,
                    mid.Z);

                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestProjectedPoint = projectedPoint;
                }
            }

            if (bestProjectedPoint == null)
                return null;

            return new XYZ(
                0.5 * (mid.X + bestProjectedPoint.X),
                0.5 * (mid.Y + bestProjectedPoint.Y),
                mid.Z);
        }

        private static bool HasParallelProjectionOverlap(Line a, Line b, double minOverlap)
        {
            if (a == null || b == null)
                return false;

            XYZ a0 = a.GetEndPoint(0);
            XYZ a1 = a.GetEndPoint(1);
            XYZ b0 = b.GetEndPoint(0);
            XYZ b1 = b.GetEndPoint(1);
            XYZ dir = Normalize2D(a1 - a0);
            if (dir == null)
                return false;

            double aStart = 0;
            double aEnd = Dot2D(a1 - a0, dir);
            if (aEnd < aStart)
            {
                double temp = aStart;
                aStart = aEnd;
                aEnd = temp;
            }

            double bStart = Dot2D(b0 - a0, dir);
            double bEnd = Dot2D(b1 - a0, dir);
            if (bEnd < bStart)
            {
                double temp = bStart;
                bStart = bEnd;
                bEnd = temp;
            }

            double overlap = Math.Min(aEnd, bEnd) - Math.Max(aStart, bStart);
            return overlap >= minOverlap;
        }

        private static bool IsLoggiaRoomName(string roomName)
        {
            if (string.IsNullOrWhiteSpace(roomName))
                return false;

            string normalized = roomName
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            return normalized.IndexOf("лоджия", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalized.IndexOf("loggia", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasRoomAtPoint(Document doc, Level level, XYZ point)
        {
            if (doc == null || level == null || point == null)
                return false;

            List<Room> rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(x => x != null && x.LevelId == level.Id && x.Area > 0)
                .ToList();

            foreach (Room room in rooms)
            {
                try
                {
                    if (room.IsPointInRoom(point))
                        return true;
                }
                catch
                {
                }
            }

            return false;
        }

        private static bool ShouldReportRoomAreaMismatch(PreparedRoomPlacement preparedRoom,
            double expectedAreaSquareMeters, double actualAreaSquareMeters, double areaToleranceSquareMeters)
        {
            if (Math.Abs(expectedAreaSquareMeters - actualAreaSquareMeters) <= areaToleranceSquareMeters)
                return false;

            if (preparedRoom != null &&
                preparedRoom.HasShaftInside &&
                actualAreaSquareMeters < expectedAreaSquareMeters)
            {
                return false;
            }

            return true;
        }

        private int PlaceRoomsForApartment(Document doc, PreparedApartmentRooms apartmentRooms, Level roomLevel,
            List<RoomAreaMismatchInfo> roomAreaMismatches, List<ElementId> createdRoomIds = null)
        {
            if (doc == null || apartmentRooms == null || roomLevel == null)
                return 0;

            if (apartmentRooms.Rooms == null || apartmentRooms.Rooms.Count == 0)
                return 0;

            int createdCount = 0;
            double areaToleranceSquareMeters = 0.1;

            foreach (PreparedRoomPlacement preparedRoom in apartmentRooms.Rooms)
            {
                if (preparedRoom == null || preparedRoom.InsertPoint == null)
                    continue;

                try
                {
                    XYZ roomPoint = new XYZ(
                        preparedRoom.InsertPoint.X,
                        preparedRoom.InsertPoint.Y,
                        roomLevel.Elevation + IDHelper.ConvertMmToInternal(100));

                    if (!apartmentRooms.HasRoomSeparators && HasRoomAtPoint(doc, roomLevel, roomPoint))
                        continue;

                    UV roomUv = new UV(preparedRoom.InsertPoint.X, preparedRoom.InsertPoint.Y);

                    Room createdRoom = doc.Create.NewRoom(roomLevel, roomUv);
                    if (createdRoom == null)
                        continue;

                    Parameter roomNameParam = createdRoom.get_Parameter(BuiltInParameter.ROOM_NAME);
                    if (roomNameParam != null && !roomNameParam.IsReadOnly && !string.IsNullOrWhiteSpace(preparedRoom.RoomName))
                        roomNameParam.Set(preparedRoom.RoomName);

                    doc.Regenerate();

                    if (preparedRoom.ExpectedAreaInternal > 0)
                    {
                        double actualAreaInternal = createdRoom.Area;
                        double expectedAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(preparedRoom.ExpectedAreaInternal);
                        double actualAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(actualAreaInternal);

                        if (ShouldReportRoomAreaMismatch(
                            preparedRoom,
                            expectedAreaSquareMeters,
                            actualAreaSquareMeters,
                            areaToleranceSquareMeters))
                        {
                            if (roomAreaMismatches != null)
                            {
                                roomAreaMismatches.Add(new RoomAreaMismatchInfo
                                {
                                    ApartmentId = preparedRoom.ApartmentId,
                                    RoomName = preparedRoom.RoomName,
                                    ExpectedAreaInternal = preparedRoom.ExpectedAreaInternal,
                                    ActualAreaInternal = actualAreaInternal
                                });
                            }
                        }
                    }

                    createdCount++;

                    if (createdRoomIds != null)
                        createdRoomIds.Add(createdRoom.Id);
                }
                catch
                {
                }
            }

            return createdCount;
        }

        private class RoomBoundaryLoopData
        {
            public CurveLoop Loop { get; set; }
            public List<XYZ> Vertices { get; set; }
            public double AreaAbs { get; set; }
        }

        private class RoomBoundaryNode
        {
            public int Id { get; set; }
            public XYZ Point { get; set; }
        }

        private class RoomBoundarySegment
        {
            public int Id { get; set; }
            public int A { get; set; }
            public int B { get; set; }
            public bool Removed { get; set; }
        }

        private static CurveLoop BuildRoomLoopFromInstance(FamilyInstance roomFi, View geometryView, List<string> debugMessages = null)
        {
            if (roomFi == null)
                throw new ArgumentNullException("roomFi");

            RoomBoundaryLoopData loopData;
            string diagnostic;
            if (TryBuildRoomBoundaryLoopFromInstanceGeometry(roomFi, geometryView, out loopData, out diagnostic))
                return loopData.Loop;

            if (debugMessages != null)
            {
                debugMessages.Add(
                    "Для вложенного помещения ID = " +
                    IDHelper.ElIdValue(roomFi.Id) +
                    " не удалось собрать замкнутый контур из геометрии экземпляра. Стены по этому помещению не создаются." +
                    (string.IsNullOrWhiteSpace(diagnostic) ? "" : " Диагностика: " + diagnostic + "."));
            }

            return null;
        }

        private static XYZ GetRoomPlacementPointFromInstance(FamilyInstance roomFi, View geometryView = null)
        {
            if (roomFi == null)
                return null;

            RoomBoundaryLoopData loopData;
            string diagnostic;
            if (TryBuildRoomBoundaryLoopFromInstanceGeometry(roomFi, geometryView, out loopData, out diagnostic))
            {
                XYZ interiorPoint = FindInteriorPointForRoomLoop(loopData.Vertices);
                if (interiorPoint != null)
                    return interiorPoint;
            }

            return null;
        }

        private static XYZ GetClosestPointOnRoomGeometry(FamilyInstance roomFi, XYZ worldPoint)
        {
            if (roomFi == null || worldPoint == null)
                return null;

            RoomBoundaryLoopData loopData;
            string diagnostic;
            if (TryBuildRoomBoundaryLoopFromInstanceGeometry(roomFi, null, out loopData, out diagnostic) &&
                loopData != null &&
                loopData.Vertices != null &&
                loopData.Vertices.Count >= 3)
            {
                if (IsPointInsidePolygon2D(worldPoint, loopData.Vertices, IDHelper.ConvertMmToInternal(10)))
                    return worldPoint;

                return GetClosestPointOnPolygonBoundary2D(worldPoint, loopData.Vertices);
            }

            return null;
        }

        private static bool IsPointInsideRoomGeometry2D(FamilyInstance roomFi, XYZ worldPoint, double tolerance, View geometryView = null)
        {
            if (roomFi == null || worldPoint == null)
                return false;

            RoomBoundaryLoopData loopData;
            string diagnostic;
            if (TryBuildRoomBoundaryLoopFromInstanceGeometry(roomFi, geometryView, out loopData, out diagnostic) &&
                loopData != null &&
                loopData.Vertices != null &&
                loopData.Vertices.Count >= 3)
            {
                return IsPointInsidePolygon2D(worldPoint, loopData.Vertices, tolerance);
            }

            return false;
        }

        private static bool TryBuildRoomBoundaryLoopFromInstanceGeometry(FamilyInstance roomFi, View geometryView, out RoomBoundaryLoopData loopData, out string diagnostic)
        {
            loopData = null;
            diagnostic = null;

            if (roomFi == null)
                return false;

            List<string> attempts = new List<string>();
            bool[] includeNonVisibleOptions = new bool[] { false, true };
            bool[] viewOptions = geometryView != null ? new bool[] { true, false } : new bool[] { false };

            foreach (bool useView in viewOptions)
            {
                View currentView = useView ? geometryView : null;

                foreach (bool includeNonVisible in includeNonVisibleOptions)
                {
                    List<Line> lines = CollectRoomBoundaryLineCandidates(roomFi, currentView, includeNonVisible);
                    List<RoomBoundaryLoopData> loops = BuildRoomBoundaryLoopsFromLines(lines);

                    attempts.Add(
                        (useView ? "вид" : "без вида") +
                        ", " +
                        (includeNonVisible ? "с невидимыми" : "видимые") +
                        ": линий " +
                        (lines != null ? lines.Count : 0) +
                        ", контуров " +
                        (loops != null ? loops.Count : 0));

                    if (loops.Count == 0)
                        continue;

                    loopData = loops
                        .OrderByDescending(x => x.AreaAbs)
                        .FirstOrDefault();

                    if (loopData != null && loopData.Loop != null)
                    {
                        diagnostic = string.Join("; ", attempts);
                        return true;
                    }
                }
            }

            diagnostic = string.Join("; ", attempts);
            return false;
        }

        private static List<Line> CollectRoomBoundaryLineCandidates(FamilyInstance roomFi, View geometryView, bool includeNonVisible)
        {
            List<Line> result = new List<Line>();
            if (roomFi == null)
                return result;

            Options options = new Options
            {
                IncludeNonVisibleObjects = includeNonVisible
            };

            if (geometryView != null)
                options.View = geometryView;
            else
                options.DetailLevel = ViewDetailLevel.Fine;

            GeometryElement geometry = null;
            try
            {
                geometry = roomFi.get_Geometry(options);
            }
            catch
            {
                geometry = null;
            }

            CollectRoomBoundaryLineCandidatesFromGeometry(geometry, Transform.Identity, result);
            return result;
        }

        private static void CollectRoomBoundaryLineCandidatesFromGeometry(GeometryElement geometry, Transform transform, List<Line> result)
        {
            if (geometry == null || transform == null || result == null)
                return;

            double minLength = IDHelper.ConvertMmToInternal(10);

            foreach (GeometryObject obj in geometry)
            {
                if (obj == null)
                    continue;

                GeometryInstance geometryInstance = obj as GeometryInstance;
                if (geometryInstance != null)
                {
                    Transform nestedTransform = transform;
                    try
                    {
                        if (geometryInstance.Transform != null)
                            nestedTransform = transform.Multiply(geometryInstance.Transform);
                    }
                    catch
                    {
                        nestedTransform = transform;
                    }

                    GeometryElement symbolGeometry = null;
                    try
                    {
                        symbolGeometry = geometryInstance.GetSymbolGeometry();
                    }
                    catch
                    {
                        symbolGeometry = null;
                    }

                    CollectRoomBoundaryLineCandidatesFromGeometry(symbolGeometry, nestedTransform, result);
                    continue;
                }

                Curve curve = obj as Curve;
                if (curve == null || !curve.IsBound)
                    continue;

                Line line = curve as Line;
                if (line != null)
                {
                    AddRoomBoundaryLineCandidate(line.GetEndPoint(0), line.GetEndPoint(1), transform, minLength, result);
                    continue;
                }

                IList<XYZ> tessellated = null;
                try
                {
                    tessellated = curve.Tessellate();
                }
                catch
                {
                    tessellated = null;
                }

                if (tessellated == null || tessellated.Count < 2)
                    continue;

                for (int i = 0; i < tessellated.Count - 1; i++)
                    AddRoomBoundaryLineCandidate(tessellated[i], tessellated[i + 1], transform, minLength, result);
            }
        }

        private static void AddRoomBoundaryLineCandidate(XYZ localP0, XYZ localP1, Transform transform, double minLength, List<Line> result)
        {
            if (localP0 == null || localP1 == null || transform == null || result == null)
                return;

            XYZ p0;
            XYZ p1;

            try
            {
                p0 = transform.OfPoint(localP0);
                p1 = transform.OfPoint(localP1);
            }
            catch
            {
                return;
            }

            if (p0 == null || p1 == null || Distance2D(p0, p1) < minLength)
                return;

            result.Add(Line.CreateBound(p0, p1));
        }

        private static List<RoomBoundaryLoopData> BuildRoomBoundaryLoopsFromLines(List<Line> lines)
        {
            List<RoomBoundaryLoopData> result = new List<RoomBoundaryLoopData>();
            if (lines == null || lines.Count < 3)
                return result;

            double nodeTol = IDHelper.ConvertMmToInternal(20);
            double minArea = IDHelper.ConvertMmToInternal(100) * IDHelper.ConvertMmToInternal(100);
            List<RoomBoundaryNode> nodes = new List<RoomBoundaryNode>();
            List<RoomBoundarySegment> segments = new List<RoomBoundarySegment>();
            HashSet<string> segmentKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Line line in lines)
            {
                if (line == null || line.Length < nodeTol)
                    continue;

                int a = GetOrCreateRoomBoundaryNode(nodes, line.GetEndPoint(0), nodeTol);
                int b = GetOrCreateRoomBoundaryNode(nodes, line.GetEndPoint(1), nodeTol);
                if (a == b)
                    continue;

                string key = a < b ? a + "|" + b : b + "|" + a;
                if (segmentKeys.Contains(key))
                    continue;

                segmentKeys.Add(key);
                segments.Add(new RoomBoundarySegment
                {
                    Id = segments.Count,
                    A = a,
                    B = b
                });
            }

            if (nodes.Count < 3 || segments.Count < 3)
                return result;

            RemoveDanglingRoomBoundarySegments(nodes, segments);

            List<List<int>> components = GetRoomBoundarySegmentComponents(nodes, segments);
            foreach (List<int> component in components)
            {
                List<XYZ> vertices;
                if (!TryOrderRoomBoundaryCycle(nodes, segments, component, out vertices))
                    continue;

                if (vertices == null || vertices.Count < 3)
                    continue;

                double area = Math.Abs(GetSignedAreaXY(vertices));
                if (area < minArea)
                    continue;

                CurveLoop loop = BuildCurveLoopFromRoomVertices(vertices);
                if (loop == null)
                    continue;

                result.Add(new RoomBoundaryLoopData
                {
                    Loop = loop,
                    Vertices = vertices,
                    AreaAbs = area
                });
            }

            return result;
        }

        private static int GetOrCreateRoomBoundaryNode(List<RoomBoundaryNode> nodes, XYZ point, double tol)
        {
            for (int i = 0; i < nodes.Count; i++)
            {
                if (Distance2D(nodes[i].Point, point) <= tol)
                    return nodes[i].Id;
            }

            int id = nodes.Count;
            nodes.Add(new RoomBoundaryNode
            {
                Id = id,
                Point = point
            });

            return id;
        }

        private static void RemoveDanglingRoomBoundarySegments(List<RoomBoundaryNode> nodes, List<RoomBoundarySegment> segments)
        {
            if (nodes == null || segments == null)
                return;

            bool changed = true;
            while (changed)
            {
                changed = false;
                Dictionary<int, int> degrees = BuildRoomBoundaryNodeDegrees(segments);

                foreach (RoomBoundarySegment segment in segments)
                {
                    if (segment == null || segment.Removed)
                        continue;

                    int degreeA = degrees.ContainsKey(segment.A) ? degrees[segment.A] : 0;
                    int degreeB = degrees.ContainsKey(segment.B) ? degrees[segment.B] : 0;
                    if (degreeA <= 1 || degreeB <= 1)
                    {
                        segment.Removed = true;
                        changed = true;
                    }
                }
            }
        }

        private static Dictionary<int, int> BuildRoomBoundaryNodeDegrees(List<RoomBoundarySegment> segments)
        {
            Dictionary<int, int> result = new Dictionary<int, int>();
            if (segments == null)
                return result;

            foreach (RoomBoundarySegment segment in segments)
            {
                if (segment == null || segment.Removed)
                    continue;

                if (!result.ContainsKey(segment.A))
                    result[segment.A] = 0;
                if (!result.ContainsKey(segment.B))
                    result[segment.B] = 0;

                result[segment.A]++;
                result[segment.B]++;
            }

            return result;
        }

        private static List<List<int>> GetRoomBoundarySegmentComponents(List<RoomBoundaryNode> nodes, List<RoomBoundarySegment> segments)
        {
            List<List<int>> result = new List<List<int>>();
            if (nodes == null || segments == null)
                return result;

            Dictionary<int, List<int>> nodeSegments = BuildRoomBoundaryNodeSegments(segments);
            HashSet<int> visitedSegments = new HashSet<int>();

            foreach (RoomBoundarySegment startSegment in segments)
            {
                if (startSegment == null || startSegment.Removed || visitedSegments.Contains(startSegment.Id))
                    continue;

                List<int> component = new List<int>();
                Queue<int> queue = new Queue<int>();
                queue.Enqueue(startSegment.Id);
                visitedSegments.Add(startSegment.Id);

                while (queue.Count > 0)
                {
                    int segmentId = queue.Dequeue();
                    component.Add(segmentId);

                    RoomBoundarySegment segment = segments[segmentId];
                    int[] nodeIds = new int[] { segment.A, segment.B };
                    foreach (int nodeId in nodeIds)
                    {
                        List<int> attached;
                        if (!nodeSegments.TryGetValue(nodeId, out attached))
                            continue;

                        foreach (int nextSegmentId in attached)
                        {
                            if (visitedSegments.Contains(nextSegmentId))
                                continue;

                            RoomBoundarySegment nextSegment = segments[nextSegmentId];
                            if (nextSegment == null || nextSegment.Removed)
                                continue;

                            visitedSegments.Add(nextSegmentId);
                            queue.Enqueue(nextSegmentId);
                        }
                    }
                }

                if (component.Count >= 3)
                    result.Add(component);
            }

            return result;
        }

        private static Dictionary<int, List<int>> BuildRoomBoundaryNodeSegments(List<RoomBoundarySegment> segments)
        {
            Dictionary<int, List<int>> result = new Dictionary<int, List<int>>();
            if (segments == null)
                return result;

            foreach (RoomBoundarySegment segment in segments)
            {
                if (segment == null || segment.Removed)
                    continue;

                AddRoomBoundaryNodeSegment(result, segment.A, segment.Id);
                AddRoomBoundaryNodeSegment(result, segment.B, segment.Id);
            }

            return result;
        }

        private static void AddRoomBoundaryNodeSegment(Dictionary<int, List<int>> nodeSegments, int nodeId, int segmentId)
        {
            List<int> attached;
            if (!nodeSegments.TryGetValue(nodeId, out attached))
            {
                attached = new List<int>();
                nodeSegments[nodeId] = attached;
            }

            attached.Add(segmentId);
        }

        private static bool TryOrderRoomBoundaryCycle(List<RoomBoundaryNode> nodes, List<RoomBoundarySegment> segments, List<int> componentSegmentIds,
            out List<XYZ> vertices)
        {
            vertices = new List<XYZ>();

            if (nodes == null || segments == null || componentSegmentIds == null || componentSegmentIds.Count < 3)
                return false;

            HashSet<int> componentSet = new HashSet<int>(componentSegmentIds);
            Dictionary<int, List<int>> nodeSegments = BuildRoomBoundaryNodeSegments(
                componentSegmentIds
                    .Select(x => segments[x])
                    .Where(x => x != null && !x.Removed)
                    .ToList());

            List<int> componentNodeIds = nodeSegments.Keys.ToList();
            if (componentNodeIds.Count < 3)
                return false;

            if (componentNodeIds.Any(x => !nodeSegments.ContainsKey(x) || nodeSegments[x].Count != 2))
                return false;

            int startNodeId = componentNodeIds
                .OrderBy(x => nodes[x].Point.X)
                .ThenBy(x => nodes[x].Point.Y)
                .First();

            int currentNodeId = startNodeId;
            int previousSegmentId = -1;
            HashSet<int> usedSegments = new HashSet<int>();

            for (int guard = 0; guard <= componentSet.Count; guard++)
            {
                vertices.Add(nodes[currentNodeId].Point);

                List<int> attached = nodeSegments[currentNodeId];
                int nextSegmentId = -1;
                foreach (int candidateSegmentId in attached)
                {
                    if (candidateSegmentId == previousSegmentId || usedSegments.Contains(candidateSegmentId))
                        continue;

                    nextSegmentId = candidateSegmentId;
                    break;
                }

                if (nextSegmentId < 0)
                    return false;

                RoomBoundarySegment nextSegment = segments[nextSegmentId];
                if (nextSegment == null || nextSegment.Removed)
                    return false;

                usedSegments.Add(nextSegmentId);

                int nextNodeId = nextSegment.A == currentNodeId ? nextSegment.B : nextSegment.A;
                if (nextNodeId == startNodeId)
                    return usedSegments.Count == componentSet.Count && vertices.Count >= 3;

                previousSegmentId = nextSegmentId;
                currentNodeId = nextNodeId;
            }

            return false;
        }

        private static CurveLoop BuildCurveLoopFromRoomVertices(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count < 3)
                return null;

            const double tol = 1e-9;
            List<Curve> profile = new List<Curve>();

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ p0 = vertices[i];
                XYZ p1 = vertices[(i + 1) % vertices.Count];

                if (p0 == null || p1 == null || p0.DistanceTo(p1) <= tol)
                    continue;

                profile.Add(Line.CreateBound(p0, p1));
            }

            if (profile.Count < 3)
                return null;

            try
            {
                return CurveLoop.Create(profile);
            }
            catch
            {
                return null;
            }
        }

        private static XYZ FindInteriorPointForRoomLoop(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count < 3)
                return null;

            XYZ centroid = TryGetPolygonCentroid(vertices);
            if (centroid != null && IsPointInsidePolygon2D(centroid, vertices, IDHelper.ConvertMmToInternal(10)))
                return centroid;

            return FindInteriorPointBySampling(vertices);
        }

        private static XYZ TryGetPolygonCentroid(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count < 3)
                return null;

            double signedArea = 0.0;
            double cx = 0.0;
            double cy = 0.0;
            double z = 0.0;

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                double cross = (a.X * b.Y) - (b.X * a.Y);
                signedArea += cross;
                cx += (a.X + b.X) * cross;
                cy += (a.Y + b.Y) * cross;
                z += a.Z;
            }

            if (Math.Abs(signedArea) < 1e-12)
                return null;

            signedArea *= 0.5;
            cx /= (6.0 * signedArea);
            cy /= (6.0 * signedArea);
            z /= vertices.Count;

            return new XYZ(cx, cy, z);
        }

        private static XYZ FindInteriorPointBySampling(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count < 3)
                return null;

            double minX = vertices.Min(x => x.X);
            double maxX = vertices.Max(x => x.X);
            double minY = vertices.Min(x => x.Y);
            double maxY = vertices.Max(x => x.Y);
            double z = vertices.Average(x => x.Z);

            double bestScore = -1.0;
            XYZ bestPoint = null;
            double tolerance = IDHelper.ConvertMmToInternal(10);
            int steps = 12;

            for (int ix = 1; ix < steps; ix++)
            {
                double x = minX + (maxX - minX) * ix / steps;

                for (int iy = 1; iy < steps; iy++)
                {
                    double y = minY + (maxY - minY) * iy / steps;
                    XYZ candidate = new XYZ(x, y, z);

                    if (!IsPointInsidePolygon2D(candidate, vertices, tolerance))
                        continue;

                    double score = DistanceToPolygonBoundary2D(candidate, vertices);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestPoint = candidate;
                    }
                }
            }

            return bestPoint;
        }

        private static bool IsPointInsidePolygon2D(XYZ point, List<XYZ> vertices, double tolerance)
        {
            if (point == null || vertices == null || vertices.Count < 3)
                return false;

            bool inside = false;
            for (int i = 0, j = vertices.Count - 1; i < vertices.Count; j = i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[j];

                if (PointOnSegment2D(point, a, b, tolerance))
                    return true;

                bool intersects =
                    ((a.Y > point.Y) != (b.Y > point.Y)) &&
                    (point.X < (b.X - a.X) * (point.Y - a.Y) / (b.Y - a.Y + 1e-30) + a.X);

                if (intersects)
                    inside = !inside;
            }

            return inside;
        }

        private static double DistanceToPolygonBoundary2D(XYZ point, List<XYZ> vertices)
        {
            if (point == null || vertices == null || vertices.Count < 2)
                return 0.0;

            double best = double.MaxValue;
            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                best = Math.Min(best, DistancePointToSegment2D(point, a, b));
            }

            return best;
        }

        private static XYZ GetClosestPointOnPolygonBoundary2D(XYZ point, List<XYZ> vertices)
        {
            if (point == null || vertices == null || vertices.Count < 2)
                return null;

            double bestDistance = double.MaxValue;
            XYZ bestPoint = null;

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                XYZ candidate = ProjectPointToSegment2D(point, a, b);
                if (candidate == null)
                    continue;

                double distance = Distance2D(point, candidate);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestPoint = candidate;
                }
            }

            return bestPoint;
        }

        private static double DistancePointToSegment2D(XYZ point, XYZ a, XYZ b)
        {
            if (point == null || a == null || b == null)
                return 0.0;

            XYZ projection = ProjectPointToSegment2D(point, a, b);
            return projection != null ? Distance2D(point, projection) : 0.0;
        }

        private static XYZ ProjectPointToSegment2D(XYZ point, XYZ a, XYZ b)
        {
            if (point == null || a == null || b == null)
                return null;

            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            double len2 = dx * dx + dy * dy;
            if (len2 < 1e-12)
                return a;

            double t = ((point.X - a.X) * dx + (point.Y - a.Y) * dy) / len2;
            t = Math.Max(0.0, Math.Min(1.0, t));

            return new XYZ(a.X + dx * t, a.Y + dy * t, point.Z);
        }

    }
}