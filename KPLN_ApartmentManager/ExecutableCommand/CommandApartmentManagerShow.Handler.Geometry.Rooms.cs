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
            Dictionary<long, ApartmentProcessState> apartmentStates, ViewPlan roomPlan, List<string> debugMessages,
            ApartmentWorksetTargets worksetTargets)
        {
            if (doc == null || preparedApartments == null || preparedApartments.Count == 0 || roomLevel == null)
                return 0;

            int createdRoomsCount = 0;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Разделители помещений"))
            {
                t.Start();
                ApplyApartmentFailureHandling(t);

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null)
                        continue;

                    ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);

                    int createdRoomSeparators = PlaceRoomSeparatorsForApartment(
                        doc,
                        apartmentWalls,
                        apartmentWalls.RoomSeparatorLines,
                        roomLevel,
                        roomPlan,
                        state,
                        debugMessages,
                        worksetTargets);
                    if (createdRoomSeparators > 0)
                    {
                        state.CreatedRoomSeparatorsCount += createdRoomSeparators;
                    }
                }

                doc.Regenerate();
                t.Commit();
            }

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

                    if (apartmentRooms != null && apartmentRooms.Rooms != null && apartmentRooms.Rooms.Count > 0)
                    {
                        createdRoomsForApartment = PlaceRoomsForApartment(
                            doc,
                            apartmentRooms,
                            roomLevel,
                            roomAreaMismatches,
                            createdRoomIds,
                            state,
                            debugMessages,
                            worksetTargets);
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
            Level roomLevel, ViewPlan roomPlan, ApartmentProcessState state, List<string> debugMessages,
            ApartmentWorksetTargets worksetTargets)
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
                TryAssignElementToWorkset(sketchPlane, worksetTargets != null ? worksetTargets.RoomWorksetId : null);
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

                        TryAssignElementToWorkset(modelCurve, worksetTargets != null ? worksetTargets.RoomWorksetId : null);
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
            List<Line> roomPlacementReferenceAxisLines, List<Line> shaftAxisLines, List<Line> roomSeparatorLines,
            PreparedApartmentDoors preparedDoors, View geometryView, List<string> debugMessages)
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
                        SourceRoom2DId = roomFi.Id,
                        RoomName = roomName.Trim(),
                        InsertPoint = insertPointInProject,
                        BoundaryVertices = loopData.Vertices != null ? loopData.Vertices.ToList() : null,
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

            AddSyntheticLoggiaRoomPlacements(result, apartmentFi, doc, loggiaAxisLines, roomPlacementReferenceAxisLines, preparedDoors, debugMessages);

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
            List<Line> loggiaAxisLines, List<Line> roomPlacementReferenceAxisLines, PreparedApartmentDoors preparedDoors, List<string> debugMessages)
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

                if (TryAddLoggiaRoomBySplittingPreparedRoom(result, loggiaAxis, apartmentInteriorPoint, preparedDoors, debugMessages))
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

                XYZ roomPoint = BuildLoggiaRoomPointBetweenWalls(loggiaAxis, roomPlacementReferenceAxisLines);

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

                PreparedRoomPlacement containingRoom = FindPreparedRoomContainingPoint(roomPoint, result.Rooms);
                if (containingRoom != null)
                {
                    if (debugMessages != null)
                    {
                        debugMessages.Add(
                            "Для квартиры ID = " + IDHelper.ElIdValue(result.ApartmentId) +
                            " точка помещения лоджии " + FormatPointMm(roomPoint) +
                            " попала внутрь подготовленного помещения '" + containingRoom.RoomName +
                            "'. Лоджия не создаётся, чтобы не разместить помещение не в той области.");
                    }

                    continue;
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
                    BoundaryVertices = null,
                    ExpectedAreaInternal = 0
                });
            }
        }

        private static bool TryAddLoggiaRoomBySplittingPreparedRoom(PreparedApartmentRooms result, Line loggiaAxis,
            XYZ apartmentInteriorPoint, PreparedApartmentDoors preparedDoors, List<string> debugMessages)
        {
            if (result == null || result.Rooms == null || result.Rooms.Count == 0 || loggiaAxis == null)
                return false;

            RoomLoggiaSplitResult bestSplit = null;
            PreparedRoomPlacement bestRoom = null;
            double bestScore = double.MinValue;

            foreach (PreparedRoomPlacement room in result.Rooms)
            {
                if (room == null || IsLoggiaRoomName(room.RoomName) ||
                    room.BoundaryVertices == null || room.BoundaryVertices.Count < 3)
                {
                    continue;
                }

                RoomLoggiaSplitResult split;
                if (!TrySplitRoomByLoggiaAxis(room, loggiaAxis, apartmentInteriorPoint, preparedDoors, out split))
                    continue;

                double score = split.MainSideResolvedByDoor
                    ? split.LoggiaAreaInternal + 1000000.0
                    : split.LoggiaAreaInternal;
                if (score > bestScore)
                {
                    bestScore = score;
                    bestSplit = split;
                    bestRoom = room;
                }
            }

            if (bestRoom == null || bestSplit == null)
                return false;

            double loggiaExpectedAreaInternal = 0;
            if (bestRoom.ExpectedAreaInternal > 0)
            {
                double splitAreaInternal = bestSplit.MainAreaInternal + bestSplit.LoggiaAreaInternal;
                if (splitAreaInternal > 1e-9 && bestSplit.MainAreaInternal > 1e-9 && bestSplit.LoggiaAreaInternal > 1e-9)
                {
                    double sourceExpectedAreaInternal = bestRoom.ExpectedAreaInternal;
                    double mainExpectedAreaInternal = sourceExpectedAreaInternal * bestSplit.MainAreaInternal / splitAreaInternal;

                    bestRoom.ExpectedAreaInternal = mainExpectedAreaInternal;
                    loggiaExpectedAreaInternal = Math.Max(0, sourceExpectedAreaInternal - mainExpectedAreaInternal);
                }
            }

            bestRoom.AreaMismatchToleranceSquareMeters = Math.Max(
                bestRoom.AreaMismatchToleranceSquareMeters,
                GetLoggiaSplitAreaToleranceSquareMeters(bestRoom.ExpectedAreaInternal));

            bestRoom.BoundaryVertices = bestSplit.MainVertices;
            bestRoom.InsertPoint = bestSplit.MainPoint;

            result.Rooms.Add(new PreparedRoomPlacement
            {
                ApartmentId = result.ApartmentId,
                RoomName = "Лоджия",
                InsertPoint = bestSplit.LoggiaPoint,
                BoundaryVertices = bestSplit.LoggiaVertices,
                ExpectedAreaInternal = loggiaExpectedAreaInternal,
                AreaMismatchToleranceSquareMeters = GetLoggiaSplitAreaToleranceSquareMeters(loggiaExpectedAreaInternal)
            });

            return true;
        }

        private static double GetLoggiaSplitAreaToleranceSquareMeters(double expectedAreaInternal)
        {
            if (expectedAreaInternal <= 0)
                return 0.3;

            double expectedAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(expectedAreaInternal);
            return Math.Max(0.3, expectedAreaSquareMeters * 0.02);
        }

        private static bool TrySplitRoomByLoggiaAxis(PreparedRoomPlacement room, Line loggiaAxis, XYZ apartmentInteriorPoint,
            PreparedApartmentDoors preparedDoors, out RoomLoggiaSplitResult result)
        {
            result = null;

            if (room == null || room.BoundaryVertices == null || room.BoundaryVertices.Count < 3 || loggiaAxis == null)
                return false;

            XYZ axisP0 = loggiaAxis.GetEndPoint(0);
            XYZ axisP1 = loggiaAxis.GetEndPoint(1);
            XYZ axisDir = Normalize2D(axisP1 - axisP0);
            if (axisDir == null)
                return false;

            double axisLength = Distance2D(axisP0, axisP1);
            if (axisLength < IDHelper.ConvertMmToInternal(10))
                return false;

            double chordFrom;
            double chordTo;
            if (!TryGetPolygonLineChordInterval(room.BoundaryVertices, axisP0, axisDir, out chordFrom, out chordTo))
                return false;

            double overlapFrom = Math.Max(0, chordFrom);
            double overlapTo = Math.Min(axisLength, chordTo);
            double minOverlap = Math.Min(axisLength * 0.35, IDHelper.ConvertMmToInternal(500));
            minOverlap = Math.Max(minOverlap, IDHelper.ConvertMmToInternal(100));
            if (overlapTo - overlapFrom < minOverlap)
                return false;

            XYZ normal = new XYZ(-axisDir.Y, axisDir.X, 0);
            double clipTol = IDHelper.ConvertMmToInternal(5);
            List<XYZ> positive = ClipPolygonByLineHalfPlane(room.BoundaryVertices, axisP0, normal, true, clipTol);
            List<XYZ> negative = ClipPolygonByLineHalfPlane(room.BoundaryVertices, axisP0, normal, false, clipTol);

            double positiveArea = GetValidPolygonArea(positive);
            double negativeArea = GetValidPolygonArea(negative);
            double minArea = IDHelper.ConvertMmToInternal(300) * IDHelper.ConvertMmToInternal(300);

            if (positiveArea < minArea || negativeArea < minArea)
                return false;

            List<XYZ> mainVertices;
            List<XYZ> loggiaVertices;
            double mainArea;
            double loggiaArea;

            double referenceSide = GetRoomMainSideByDoorPoints(room, preparedDoors, axisP0, normal, clipTol);
            bool mainSideResolvedByDoor = Math.Abs(referenceSide) > clipTol;

            if (Math.Abs(referenceSide) <= clipTol && room.InsertPoint != null)
                referenceSide = Dot2D(room.InsertPoint - axisP0, normal);

            if (Math.Abs(referenceSide) <= clipTol && apartmentInteriorPoint != null)
                referenceSide = Dot2D(apartmentInteriorPoint - axisP0, normal);

            if (Math.Abs(referenceSide) > clipTol)
            {
                bool mainIsPositive = referenceSide > 0;
                mainVertices = mainIsPositive ? positive : negative;
                loggiaVertices = mainIsPositive ? negative : positive;
                mainArea = mainIsPositive ? positiveArea : negativeArea;
                loggiaArea = mainIsPositive ? negativeArea : positiveArea;
            }
            else if (positiveArea >= negativeArea)
            {
                mainVertices = positive;
                loggiaVertices = negative;
                mainArea = positiveArea;
                loggiaArea = negativeArea;
            }
            else
            {
                mainVertices = negative;
                loggiaVertices = positive;
                mainArea = negativeArea;
                loggiaArea = positiveArea;
            }

            XYZ mainPoint = FindInteriorPointForRoomLoop(mainVertices);
            XYZ loggiaPoint = FindInteriorPointForRoomLoop(loggiaVertices);
            if (mainPoint == null || loggiaPoint == null)
                return false;

            result = new RoomLoggiaSplitResult
            {
                MainVertices = mainVertices,
                LoggiaVertices = loggiaVertices,
                MainPoint = mainPoint,
                LoggiaPoint = loggiaPoint,
                MainAreaInternal = mainArea,
                LoggiaAreaInternal = loggiaArea,
                MainSideResolvedByDoor = mainSideResolvedByDoor
            };

            return true;
        }

        private static double GetRoomMainSideByDoorPoints(PreparedRoomPlacement room, PreparedApartmentDoors preparedDoors,
            XYZ linePoint, XYZ normal, double tolerance)
        {
            if (room == null || preparedDoors == null || preparedDoors.Doors == null ||
                linePoint == null || normal == null)
            {
                return 0;
            }

            int positiveCount = 0;
            int negativeCount = 0;

            foreach (PreparedDoorPlacement door in preparedDoors.Doors)
            {
                if (door == null || door.InsertPoint == null || door.IsEntranceDoor)
                    continue;

                if (!DoorBelongsToPreparedRoom(door, room) && !DoorPointBelongsToPreparedRoomGeometry(door, room))
                    continue;

                double side = Dot2D(door.InsertPoint - linePoint, normal);
                if (side > tolerance)
                    positiveCount++;
                else if (side < -tolerance)
                    negativeCount++;
            }

            if (positiveCount == 0 && negativeCount == 0)
                return 0;

            if (positiveCount == negativeCount)
                return 0;

            if (positiveCount > negativeCount)
                return 1;

            return -1;
        }

        private static bool DoorBelongsToPreparedRoom(PreparedDoorPlacement door, PreparedRoomPlacement room)
        {
            if (door == null || room == null)
                return false;

            if (room.SourceRoom2DId != null && room.SourceRoom2DId != ElementId.InvalidElementId &&
                door.RelatedRoom2D != null && door.RelatedRoom2D.Id != null &&
                IDHelper.ElIdValue(door.RelatedRoom2D.Id) == IDHelper.ElIdValue(room.SourceRoom2DId))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(door.RoomCategory) &&
                !string.Equals(door.RoomCategory, "-", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(room.RoomName) &&
                string.Equals(door.RoomCategory.Trim(), room.RoomName.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static bool DoorPointBelongsToPreparedRoomGeometry(PreparedDoorPlacement door, PreparedRoomPlacement room)
        {
            if (door == null || door.InsertPoint == null || room == null ||
                room.BoundaryVertices == null || room.BoundaryVertices.Count < 3)
            {
                return false;
            }

            double tolerance = IDHelper.ConvertMmToInternal(150);
            return IsPointInsidePolygon2D(door.InsertPoint, room.BoundaryVertices, tolerance);
        }

        private static bool TryGetPolygonLineChordInterval(List<XYZ> vertices, XYZ linePoint, XYZ lineDir, out double from, out double to)
        {
            from = 0;
            to = 0;

            if (vertices == null || vertices.Count < 3 || linePoint == null || lineDir == null)
                return false;

            XYZ normal = new XYZ(-lineDir.Y, lineDir.X, 0);
            double tol = IDHelper.ConvertMmToInternal(5);
            List<double> parameters = new List<double>();

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                if (a == null || b == null)
                    continue;

                double da = Dot2D(a - linePoint, normal);
                double db = Dot2D(b - linePoint, normal);

                if (Math.Abs(da) <= tol)
                    AddDistinctLineParameter(parameters, Dot2D(a - linePoint, lineDir), tol);

                if (Math.Abs(db) <= tol)
                    AddDistinctLineParameter(parameters, Dot2D(b - linePoint, lineDir), tol);

                if ((da > tol && db < -tol) || (da < -tol && db > tol))
                {
                    double ratio = da / (da - db);
                    XYZ intersection = new XYZ(
                        a.X + (b.X - a.X) * ratio,
                        a.Y + (b.Y - a.Y) * ratio,
                        a.Z + (b.Z - a.Z) * ratio);

                    AddDistinctLineParameter(parameters, Dot2D(intersection - linePoint, lineDir), tol);
                }
            }

            if (parameters.Count < 2)
                return false;

            List<double> ordered = parameters.OrderBy(x => x).ToList();
            from = ordered.First();
            to = ordered.Last();

            return to - from >= IDHelper.ConvertMmToInternal(100);
        }

        private static void AddDistinctLineParameter(List<double> parameters, double value, double tolerance)
        {
            if (parameters == null)
                return;

            foreach (double existing in parameters)
            {
                if (Math.Abs(existing - value) <= tolerance)
                    return;
            }

            parameters.Add(value);
        }

        private static List<XYZ> ClipPolygonByLineHalfPlane(List<XYZ> vertices, XYZ linePoint, XYZ normal, bool keepPositive, double tolerance)
        {
            List<XYZ> result = new List<XYZ>();
            if (vertices == null || vertices.Count < 3 || linePoint == null || normal == null)
                return result;

            XYZ previous = vertices[vertices.Count - 1];
            double previousDistance = Dot2D(previous - linePoint, normal);
            bool previousInside = keepPositive
                ? previousDistance >= -tolerance
                : previousDistance <= tolerance;

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ current = vertices[i];
                if (current == null || previous == null)
                {
                    previous = current;
                    continue;
                }

                double currentDistance = Dot2D(current - linePoint, normal);
                bool currentInside = keepPositive
                    ? currentDistance >= -tolerance
                    : currentDistance <= tolerance;

                if (previousInside != currentInside)
                {
                    XYZ intersection = IntersectSegmentWithLine(previous, current, linePoint, normal);
                    AddDistinctPolygonVertex(result, intersection, tolerance);
                }

                if (currentInside)
                    AddDistinctPolygonVertex(result, current, tolerance);

                previous = current;
                previousDistance = currentDistance;
                previousInside = currentInside;
            }

            return NormalizePolygonVertices(result, tolerance);
        }

        private static XYZ IntersectSegmentWithLine(XYZ a, XYZ b, XYZ linePoint, XYZ normal)
        {
            double da = Dot2D(a - linePoint, normal);
            double db = Dot2D(b - linePoint, normal);
            double denom = da - db;
            if (Math.Abs(denom) < 1e-12)
                return a;

            double ratio = da / denom;
            if (ratio < 0)
                ratio = 0;
            else if (ratio > 1)
                ratio = 1;

            return new XYZ(
                a.X + (b.X - a.X) * ratio,
                a.Y + (b.Y - a.Y) * ratio,
                a.Z + (b.Z - a.Z) * ratio);
        }

        private static void AddDistinctPolygonVertex(List<XYZ> vertices, XYZ point, double tolerance)
        {
            if (vertices == null || point == null)
                return;

            if (vertices.Count > 0 && Distance2D(vertices[vertices.Count - 1], point) <= tolerance)
                return;

            vertices.Add(point);
        }

        private static List<XYZ> NormalizePolygonVertices(List<XYZ> vertices, double tolerance)
        {
            List<XYZ> result = new List<XYZ>();
            if (vertices == null)
                return result;

            foreach (XYZ vertex in vertices)
                AddDistinctPolygonVertex(result, vertex, tolerance);

            while (result.Count > 1 && Distance2D(result[0], result[result.Count - 1]) <= tolerance)
                result.RemoveAt(result.Count - 1);

            return result;
        }

        private static double GetValidPolygonArea(List<XYZ> vertices)
        {
            if (vertices == null || vertices.Count < 3)
                return 0;

            return Math.Abs(GetSignedAreaXY(vertices));
        }

        private static PreparedRoomPlacement FindPreparedRoomContainingPoint(XYZ point, List<PreparedRoomPlacement> rooms)
        {
            if (point == null || rooms == null || rooms.Count == 0)
                return null;

            double tolerance = IDHelper.ConvertMmToInternal(10);

            foreach (PreparedRoomPlacement room in rooms)
            {
                if (room == null || IsLoggiaRoomName(room.RoomName) ||
                    room.BoundaryVertices == null || room.BoundaryVertices.Count < 3)
                {
                    continue;
                }

                if (IsPointInsidePolygon2D(point, room.BoundaryVertices, tolerance))
                    return room;
            }

            return null;
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

        private static bool HasRoomAtPoint(Document doc, Level level, XYZ point, IEnumerable<ElementId> excludedRoomIds = null)
        {
            if (doc == null || level == null || point == null)
                return false;

            HashSet<long> excludedIds = new HashSet<long>();
            if (excludedRoomIds != null)
            {
                foreach (ElementId id in excludedRoomIds)
                {
                    if (id != null && id != ElementId.InvalidElementId)
                        excludedIds.Add(IDHelper.ElIdValue(id));
                }
            }

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
                    if (room.Id != null && excludedIds.Contains(IDHelper.ElIdValue(room.Id)))
                        continue;

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

        private static void AddRoomAreaMismatchIfNeeded(PreparedRoomPlacement preparedRoom, double actualAreaInternal,
            List<RoomAreaMismatchInfo> roomAreaMismatches, double areaToleranceSquareMeters)
        {
            if (preparedRoom == null || preparedRoom.ExpectedAreaInternal <= 0 || roomAreaMismatches == null)
                return;

            double expectedAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(preparedRoom.ExpectedAreaInternal);
            double actualAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(actualAreaInternal);
            double effectiveAreaToleranceSquareMeters = Math.Max(
                areaToleranceSquareMeters,
                preparedRoom.AreaMismatchToleranceSquareMeters);

            if (!ShouldReportRoomAreaMismatch(
                preparedRoom,
                expectedAreaSquareMeters,
                actualAreaSquareMeters,
                effectiveAreaToleranceSquareMeters))
            {
                return;
            }

            roomAreaMismatches.Add(new RoomAreaMismatchInfo
            {
                ApartmentId = preparedRoom.ApartmentId,
                RoomName = preparedRoom.RoomName,
                ExpectedAreaInternal = preparedRoom.ExpectedAreaInternal,
                ActualAreaInternal = actualAreaInternal
            });
        }

        private int PlaceRoomsForApartment(Document doc, PreparedApartmentRooms apartmentRooms, Level roomLevel,
            List<RoomAreaMismatchInfo> roomAreaMismatches, List<ElementId> createdRoomIds = null,
            ApartmentProcessState state = null, List<string> debugMessages = null,
            ApartmentWorksetTargets worksetTargets = null)
        {
            if (doc == null || apartmentRooms == null || roomLevel == null)
                return 0;

            if (apartmentRooms.Rooms == null || apartmentRooms.Rooms.Count == 0)
                return 0;

            int createdCount = 0;
            double areaToleranceSquareMeters = 0.1;
            List<CreatedRoomPlacementInfo> roomsPendingSeparatorAreaCheck = apartmentRooms.HasRoomSeparators
                ? new List<CreatedRoomPlacementInfo>()
                : null;

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

                    IEnumerable<ElementId> ignoredRoomsForPointCheck = apartmentRooms.HasRoomSeparators
                        ? createdRoomIds
                        : null;

                    if (HasRoomAtPoint(doc, roomLevel, roomPoint, ignoredRoomsForPointCheck))
                    {
                        AddApartmentDiagnostic(
                            state,
                            debugMessages,
                            "Помещение '" + preparedRoom.RoomName + "' не создано: в точке " +
                            FormatPointMm(roomPoint) + " уже есть помещение на уровне '" + roomLevel.Name + "'.");
                        continue;
                    }

                    UV roomUv = new UV(preparedRoom.InsertPoint.X, preparedRoom.InsertPoint.Y);

                    Room createdRoom = doc.Create.NewRoom(roomLevel, roomUv);
                    if (createdRoom == null)
                    {
                        AddApartmentDiagnostic(
                            state,
                            debugMessages,
                            "Revit не создал помещение '" + preparedRoom.RoomName +
                            "' в точке " + FormatPointMm(roomPoint) + ".");
                        continue;
                    }

                    TryAssignElementToWorkset(createdRoom, worksetTargets != null ? worksetTargets.RoomWorksetId : null);

                    Parameter roomNameParam = createdRoom.get_Parameter(BuiltInParameter.ROOM_NAME);
                    if (roomNameParam != null && !roomNameParam.IsReadOnly && !string.IsNullOrWhiteSpace(preparedRoom.RoomName))
                        roomNameParam.Set(preparedRoom.RoomName);

                    doc.Regenerate();

                    if (apartmentRooms.HasRoomSeparators)
                    {
                        createdCount++;

                        if (createdRoomIds != null)
                            createdRoomIds.Add(createdRoom.Id);

                        roomsPendingSeparatorAreaCheck.Add(new CreatedRoomPlacementInfo
                        {
                            RoomId = createdRoom.Id,
                            PreparedRoom = preparedRoom
                        });

                        continue;
                    }

                    if (createdRoom.Area <= 1e-9)
                    {
                        ElementId createdRoomId = createdRoom.Id;
                        try
                        {
                            doc.Delete(createdRoomId);
                        }
                        catch
                        {
                        }

                        AddRoomAreaMismatchIfNeeded(
                            preparedRoom,
                            0,
                            roomAreaMismatches,
                            areaToleranceSquareMeters);

                        AddApartmentDiagnostic(
                            state,
                            debugMessages,
                            "Помещение '" + preparedRoom.RoomName + "' в точке " +
                            FormatPointMm(roomPoint) +
                            " получило нулевую площадь и удалено. Проверьте замкнутость и Room Bounding стен вокруг точки.");

                        continue;
                    }

                    AddRoomAreaMismatchIfNeeded(
                        preparedRoom,
                        createdRoom.Area,
                        roomAreaMismatches,
                        areaToleranceSquareMeters);

                    createdCount++;

                    if (createdRoomIds != null)
                        createdRoomIds.Add(createdRoom.Id);
                }
                catch (Exception ex)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Ошибка создания помещения '" + preparedRoom.RoomName + "' в точке " +
                        FormatPointMm(preparedRoom.InsertPoint) + ": " + ex.Message);
                }
            }

            if (roomsPendingSeparatorAreaCheck != null && roomsPendingSeparatorAreaCheck.Count > 0)
            {
                doc.Regenerate();

                foreach (CreatedRoomPlacementInfo createdRoomInfo in roomsPendingSeparatorAreaCheck)
                {
                    if (createdRoomInfo == null || createdRoomInfo.RoomId == null ||
                        createdRoomInfo.RoomId == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    try
                    {
                        Room createdRoom = doc.GetElement(createdRoomInfo.RoomId) as Room;
                        if (createdRoom == null)
                            continue;

                        if (createdRoom.Area <= 1e-9)
                        {
                            try
                            {
                                doc.Delete(createdRoomInfo.RoomId);
                            }
                            catch
                            {
                            }

                            if (createdRoomIds != null)
                            {
                                createdRoomIds.RemoveAll(x =>
                                    x != null &&
                                    IDHelper.ElIdValue(x) == IDHelper.ElIdValue(createdRoomInfo.RoomId));
                            }

                            createdCount = Math.Max(0, createdCount - 1);

                            AddRoomAreaMismatchIfNeeded(
                                createdRoomInfo.PreparedRoom,
                                0,
                                roomAreaMismatches,
                                areaToleranceSquareMeters);

                            continue;
                        }

                        AddRoomAreaMismatchIfNeeded(
                            createdRoomInfo.PreparedRoom,
                            createdRoom.Area,
                            roomAreaMismatches,
                            areaToleranceSquareMeters);
                    }
                    catch (Exception ex)
                    {
                        string roomName = createdRoomInfo.PreparedRoom != null
                            ? createdRoomInfo.PreparedRoom.RoomName
                            : "Помещение";

                        AddApartmentDiagnostic(
                            state,
                            debugMessages,
                            "Ошибка проверки помещения '" + roomName +
                            "' после создания разделителей: " + ex.Message);
                    }
                }
            }

            return createdCount;
        }

        private class CreatedRoomPlacementInfo
        {
            public ElementId RoomId { get; set; }
            public PreparedRoomPlacement PreparedRoom { get; set; }
        }

        private class RoomBoundaryLoopData
        {
            public CurveLoop Loop { get; set; }
            public List<XYZ> Vertices { get; set; }
            public double AreaAbs { get; set; }
        }

        private class RoomLoggiaSplitResult
        {
            public List<XYZ> MainVertices { get; set; }
            public List<XYZ> LoggiaVertices { get; set; }
            public XYZ MainPoint { get; set; }
            public XYZ LoggiaPoint { get; set; }
            public double MainAreaInternal { get; set; }
            public double LoggiaAreaInternal { get; set; }
            public bool MainSideResolvedByDoor { get; set; }
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
