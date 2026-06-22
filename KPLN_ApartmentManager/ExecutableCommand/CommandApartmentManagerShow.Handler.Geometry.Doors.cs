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
        private int PlaceDoorGeometryInTransaction(Document doc, List<PreparedApartmentWalls> preparedApartments,
            List<PreparedApartmentDoors> preparedDoorsByApartment, Dictionary<long, List<Wall>> doorHostWallsByApartment,
            List<ExistingWallLineInfo> existingWalls, Level baseLevel, List<string> debugMessages,
            Dictionary<long, ApartmentProcessState> apartmentStates, ApartmentWorksetTargets worksetTargets)
        {
            if (doc == null || preparedApartments == null || preparedApartments.Count == 0 || baseLevel == null)
                return 0;

            int installedDoorsCount = 0;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Геометрия дверей"))
            {
                t.Start();
                ApplyApartmentFailureHandling(t);

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null)
                        continue;

                    PreparedApartmentDoors apartmentDoors = preparedDoorsByApartment != null
                        ? preparedDoorsByApartment.FirstOrDefault(x => x != null && x.ApartmentId == apartmentWalls.ApartmentId)
                        : null;

                    if (apartmentDoors == null || apartmentDoors.Doors == null || apartmentDoors.Doors.Count == 0)
                        continue;

                    ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);
                    List<Wall> createdDoorHostWallsForApartment = GetCreatedWallsForApartment(doorHostWallsByApartment, apartmentWalls.ApartmentId);
                    List<ElementId> createdDoorIds = new List<ElementId>();

                    int installedDoorsForApartment = PlaceDoorsForApartment(
                        doc,
                        apartmentDoors,
                        createdDoorHostWallsForApartment,
                        existingWalls,
                        baseLevel,
                        debugMessages,
                        state,
                        createdDoorIds,
                        worksetTargets);

                    foreach (ElementId createdDoorId in createdDoorIds)
                        AddCreatedElementCandidate(state, createdDoorId);

                    installedDoorsCount += installedDoorsForApartment;

                    if (installedDoorsForApartment > 0)
                        state.HasInstalledDoors = true;

                    int skippedDoorsForApartment = apartmentDoors.Doors.Count - installedDoorsForApartment;
                    if (skippedDoorsForApartment > 0)
                        state.SkippedDoorsCount += skippedDoorsForApartment;
                }

                doc.Regenerate();
                t.Commit();
            }

            return installedDoorsCount;
        }

        private static XYZ ResolveDoorPlacementSideDirection(PreparedDoorPlacement preparedDoor, Wall hostWall, XYZ projectedPoint)
        {
            if (preparedDoor == null || hostWall == null || projectedPoint == null)
                return null;

            XYZ sourceSideDirection = Normalize2D(preparedDoor.SourceRoomCalculationSideDirection);
            if (sourceSideDirection == null)
                return null;

            XYZ wallNormal = GetWallAxisNormal2D(hostWall);
            if (wallNormal == null)
                return null;

            double side = Dot2D(sourceSideDirection, wallNormal);
            if (Math.Abs(side) < 0.15)
                return null;

            return side >= 0
                ? wallNormal
                : new XYZ(-wallNormal.X, -wallNormal.Y, 0);
        }

        private static void ApplyDoorPlacementSideDirection(Document doc, FamilyInstance createdDoor, PreparedDoorPlacement preparedDoor,
            Wall hostWall, XYZ projectedPoint, List<string> debugMessages)
        {
            XYZ desiredSideDirection = ResolveDoorPlacementSideDirection(preparedDoor, hostWall, projectedPoint);
            if (createdDoor == null || desiredSideDirection == null)
                return;

            TryRegenerateDoorDocument(doc);

            bool isOnDesiredSide;
            double currentScore;
            if (!TryIsDoorCalculationPointOnDesiredSide(createdDoor, projectedPoint, desiredSideDirection, out isOnDesiredSide, out currentScore))
                return;

            if (isOnDesiredSide)
                return;

            if (debugMessages != null)
            {
                debugMessages.Add(
                    "Дверь ID = " + IDHelper.ElIdValue(createdDoor.Id) +
                    " создана не на нужной стороне стены. Пост-коррекция через flipFacing отключена; " +
                    "дверь должна создаваться сразу с правильной точкой вставки. score = " +
                    FormatDouble(currentScore) + ".");
            }
        }

        private static void TryRegenerateDoorDocument(Document doc)
        {
            if (doc == null)
                return;

            try
            {
                doc.Regenerate();
            }
            catch
            {
            }
        }

        private static bool TryIsDoorCalculationPointOnDesiredSide(FamilyInstance door, XYZ projectedPoint, XYZ desiredSideDirection,
            out bool isOnDesiredSide, out double score)
        {
            isOnDesiredSide = false;
            score = double.NegativeInfinity;

            if (door == null || projectedPoint == null || desiredSideDirection == null)
                return false;

            XYZ calculationPoint;
            if (!TryGetDoorSpatialElementCalculationPoint(door, out calculationPoint) || calculationPoint == null)
                return false;

            XYZ calculationDirection = Normalize2D(calculationPoint - projectedPoint);
            if (calculationDirection == null)
                return false;

            score = Dot2D(calculationDirection, desiredSideDirection);
            if (Math.Abs(score) < 0.05)
                return false;

            isOnDesiredSide = score > 0;
            return true;
        }

        private static bool TryGetDoorSpatialElementCalculationPoint(FamilyInstance door, out XYZ point)
        {
            point = null;

            if (door == null || door.Document == null)
                return false;

            try
            {
                ICollection<ElementId> dependentIds = door.GetDependentElements(new ElementClassFilter(typeof(SpatialElementCalculationPoint)));
                if (dependentIds != null)
                {
                    foreach (ElementId dependentId in dependentIds)
                    {
                        SpatialElementCalculationPoint calculationPoint =
                            door.Document.GetElement(dependentId) as SpatialElementCalculationPoint;

                        if (calculationPoint == null || calculationPoint.Position == null)
                            continue;

                        point = calculationPoint.Position;
                        return true;
                    }
                }
            }
            catch
            {
                point = null;
            }

            return false;
        }

        private static XYZ GetDoorRoomCalculationSideDirection(FamilyInstance door, XYZ insertPoint)
        {
            if (door == null || insertPoint == null)
                return null;

            XYZ calculationPoint;
            if (!TryGetDoorSpatialElementCalculationPoint(door, out calculationPoint) || calculationPoint == null)
                return null;

            XYZ sideDirection = Normalize2D(calculationPoint - insertPoint);
            if (sideDirection == null)
                return null;

            return sideDirection;
        }

        private static XYZ GetRoomCenterPoint(FamilyInstance roomFi)
        {
            return GetRoomPlacementPointFromInstance(roomFi);
        }

        private static FamilyInstance FindBestMatchingRoomForDoor(
            FamilyInstance apartmentFi,
            string roomCategory,
            XYZ doorPoint,
            Document doc)
        {
            List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartmentFi);
            if (rooms == null || rooms.Count == 0)
                return null;

            IEnumerable<FamilyInstance> filteredRooms = rooms;

            if (!string.IsNullOrWhiteSpace(roomCategory) && roomCategory != "-")
            {
                List<FamilyInstance> exactRooms = rooms
                    .Where(x => string.Equals(
                        GetRoomCategoryLabel(x),
                        roomCategory,
                        StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (exactRooms.Count > 0)
                    filteredRooms = exactRooms;
            }

            FamilyInstance bestRoom = null;
            double bestDistance = double.MaxValue;

            foreach (FamilyInstance roomFi in filteredRooms)
            {
                XYZ center = GetRoomCenterPoint(roomFi);
                if (center == null)
                    continue;

                double dist = Distance2D(center, doorPoint);
                if (dist < bestDistance)
                {
                    bestDistance = dist;
                    bestRoom = roomFi;
                }
            }

            return bestRoom;
        }

        private static XYZ GetClosestPointOnRoomRectangle(FamilyInstance roomFi, XYZ worldPoint)
        {
            return GetClosestPointOnRoomGeometry(roomFi, worldPoint);
        }

        private int PlaceDoorsForApartment(Document doc, PreparedApartmentDoors apartmentDoors, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, Level baseLevel, List<string> debugMessages, ApartmentProcessState state,
            List<ElementId> createdDoorIds = null, ApartmentWorksetTargets worksetTargets = null)
        {
            if (doc == null || apartmentDoors == null || createdWallsForApartment == null || baseLevel == null)
                return 0;

            if (apartmentDoors.Doors == null || apartmentDoors.Doors.Count == 0)
                return 0;

            int installedCount = 0;
            double maxDistanceToWallAxis = IDHelper.ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                if (preparedDoor == null || preparedDoor.DoorSymbol == null || preparedDoor.InsertPoint == null)
                    continue;

                if (preparedDoor.IsEntranceDoor)
                {
                    if (PlaceEntranceDoorForApartment(
                        doc,
                        preparedDoor,
                        createdWallsForApartment,
                        existingWallsOnLevel,
                        baseLevel,
                        debugMessages,
                        state,
                        createdDoorIds,
                        worksetTargets))
                    {
                        installedCount++;
                    }

                    continue;
                }

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;
                bool hostFromExistingWall;
                string roomLookupValue;
                string roomDiagnostic;
                Room expectedRoom;
                bool hasExpectedRoomValue = TryGetSourceDoorRoom(
                    doc,
                    preparedDoor,
                    out expectedRoom,
                    out roomLookupValue,
                    out roomDiagnostic);

                bool foundHost = TryFindHostWallForDoorPlacement(
                    doc,
                    preparedDoor,
                    createdWallsForApartment,
                    existingWallsOnLevel,
                    maxDistanceToWallAxis,
                    out hostWall,
                    out projectedPoint,
                    out distanceToWallAxis,
                    out hostFromExistingWall);

                if (!hasExpectedRoomValue)
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не удалось прочитать Room у исходной 2D-двери" +
                        " ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        " квартиры ID = " + IDHelper.ElIdValue(preparedDoor.ApartmentId) + ". " +
                        roomDiagnostic);

                    continue;
                }

                if (!foundHost || hostWall == null || projectedPoint == null)
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не найдена стена-хост для двери" +
                        " ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        " квартиры ID = " + IDHelper.ElIdValue(preparedDoor.ApartmentId) + ". " +
                        BuildDoorHostSearchDiagnostic(doc, preparedDoor, existingWallsOnLevel, createdWallsForApartment, maxDistanceToWallAxis));

                    continue;
                }

                FamilySymbol symbolToPlace = preparedDoor.DoorSymbol;

                if (symbolToPlace == null)
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Для двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        " не удалось определить итоговый тип перед вставкой.");
                    continue;
                }

                projectedPoint = AlignHostedFamilyInsertionPointToHostWallBase(projectedPoint, hostWall, baseLevel);

                try
                {
                    if (!symbolToPlace.IsActive)
                    {
                        symbolToPlace.Activate();
                        doc.Regenerate();
                    }
                }
                catch (Exception ex)
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не удалось активировать тип двери '" + BuildDoorTypeDisplayName(symbolToPlace) +
                        "' для 2D-двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) + ": " + ex.Message);
                    continue;
                }

                FamilyInstance createdDoor = null;
                string attemptsDiagnostic = null;

                if (!TryCreateDoorInExpectedRoom(
                    doc,
                    preparedDoor,
                    symbolToPlace,
                    hostWall,
                    projectedPoint,
                    baseLevel,
                    expectedRoom,
                    createdDoorIds,
                    worksetTargets,
                    out createdDoor,
                    out attemptsDiagnostic))
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не найден вариант вставки двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        " с нужным Room = '" + FormatDiagnosticValue(roomLookupValue) + "'" +
                        " в стену ID = " + IDHelper.ElIdValue(hostWall.Id) +
                        ", тип '" + BuildDoorTypeDisplayName(symbolToPlace) + "'. Попытки: " +
                        attemptsDiagnostic);
                    continue;
                }

                if (createdDoor != null)
                {
                    TryAssignElementToWorkset(createdDoor, worksetTargets != null ? worksetTargets.DoorWorksetId : null);
                    installedCount++;

                    if (createdDoorIds != null)
                        createdDoorIds.Add(createdDoor.Id);
                }
                else
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Revit не создал дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                       " в стене ID = " + IDHelper.ElIdValue(hostWall.Id) +
                        ", тип '" + BuildDoorTypeDisplayName(symbolToPlace) + "'.");
                }
            }

            return installedCount;
        }

        private static bool TryGetSourceDoorRoom(Document doc, PreparedDoorPlacement preparedDoor,
            out Room room, out string roomLookupValue, out string diagnostic)
        {
            room = null;
            roomLookupValue = "<нет>";
            diagnostic = "<нет данных>";

            if (doc == null || preparedDoor == null || preparedDoor.Door2DId == null ||
                preparedDoor.Door2DId == ElementId.InvalidElementId)
            {
                diagnostic = "Нет ID исходной 2D-двери.";
                return false;
            }

            FamilyInstance sourceDoor = doc.GetElement(preparedDoor.Door2DId) as FamilyInstance;
            if (sourceDoor == null)
            {
                diagnostic = "Исходная 2D-дверь не найдена в документе.";
                return false;
            }

            try
            {
                room = sourceDoor.Room;
            }
            catch (Exception ex)
            {
                diagnostic = "Свойство FamilyInstance.Room не прочиталось: " + ex.Message;
                return false;
            }

            roomLookupValue = FormatRoomLookupValue(room);
            diagnostic = "OK";
            return true;
        }

        private static string FormatRoomLookupValue(Room room)
        {
            if (room == null)
                return "<нет>";

            List<string> parts = new List<string>();

            try
            {
                if (!string.IsNullOrWhiteSpace(room.Name))
                    parts.Add(room.Name.Trim());
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(room.Number))
                    parts.Add(room.Number.Trim());
            }
            catch
            {
            }

            try
            {
                parts.Add(IDHelper.ElIdValue(room.Id).ToString());
            }
            catch
            {
            }

            return parts.Count > 0 ? string.Join(" ", parts.ToArray()) : "<нет>";
        }

        private static string FormatDiagnosticValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<нет>" : value.Trim();
        }

        private enum DoorRoomSideMatch
        {
            None,
            FromRoom,
            ToRoom
        }

        private static bool TryCreateDoorInExpectedRoom(Document doc, PreparedDoorPlacement preparedDoor, FamilySymbol symbolToPlace,
            Wall hostWall, XYZ projectedPoint, Level baseLevel, Room expectedRoom, List<ElementId> createdElementIds,
            ApartmentWorksetTargets worksetTargets,
            out FamilyInstance createdDoor, out string diagnostic)
        {
            createdDoor = null;
            diagnostic = "<нет попыток>";

            if (doc == null || symbolToPlace == null || hostWall == null || projectedPoint == null)
            {
                diagnostic = "Нет документа, типа двери, стены-хоста или точки вставки.";
                return false;
            }

            List<string> attemptMessages = new List<string>();

            FamilyInstance initialDoor;
            DoorRoomSideMatch initialMatch;
            string initialDiagnostic;
            if (!TryCreateDoorAndReadRoomSide(
                doc,
                "исходная сторона стены",
                symbolToPlace,
                hostWall,
                projectedPoint,
                baseLevel,
                expectedRoom,
                out initialDoor,
                out initialMatch,
                out initialDiagnostic))
            {
                diagnostic = initialDiagnostic;
                return false;
            }

            attemptMessages.Add(initialDiagnostic);

            if (initialMatch == DoorRoomSideMatch.ToRoom)
            {
                createdDoor = initialDoor;
                diagnostic = string.Join("; ", attemptMessages.ToArray());
                return true;
            }

            if (initialMatch != DoorRoomSideMatch.FromRoom)
            {
                if (initialDoor != null)
                    TryDeleteElement(doc, initialDoor.Id, attemptMessages);

                diagnostic = string.Join("; ", attemptMessages.ToArray());
                return false;
            }

            if (initialDoor != null)
                TryDeleteElement(doc, initialDoor.Id, attemptMessages);

            TryRegenerateDoorDocument(doc);

            string auxiliaryDiagnostic;
            if (TryCreateDoorWithAuxiliaryHostWallForRoomSide(
                doc,
                preparedDoor,
                symbolToPlace,
                hostWall,
                projectedPoint,
                baseLevel,
                expectedRoom,
                createdElementIds,
                worksetTargets,
                out createdDoor,
                out auxiliaryDiagnostic))
            {
                attemptMessages.Add(auxiliaryDiagnostic);
                diagnostic = string.Join("; ", attemptMessages.ToArray());
                return true;
            }

            attemptMessages.Add(auxiliaryDiagnostic);
            diagnostic = string.Join("; ", attemptMessages.ToArray());
            return false;
        }

        private static bool TryCreateDoorAndReadRoomSide(Document doc, string attemptName, FamilySymbol symbolToPlace,
            Wall hostWall, XYZ projectedPoint, Level baseLevel, Room expectedRoom,
            out FamilyInstance createdDoor, out DoorRoomSideMatch match, out string diagnostic)
        {
            createdDoor = null;
            match = DoorRoomSideMatch.None;
            diagnostic = attemptName + ": <нет данных>";

            XYZ referenceDirection = GetWallAxisDirection2D(hostWall);
            bool useLevelHostOverload = ShouldCreateHostedFamilyWithLevelHostOverload(hostWall, baseLevel);
            if (!useLevelHostOverload && referenceDirection == null)
            {
                diagnostic = attemptName + ": не удалось определить вектор оси стены.";
                return false;
            }

            try
            {
                createdDoor = useLevelHostOverload
                    ? TryCreateDoorDefault(doc, projectedPoint, symbolToPlace, hostWall, baseLevel)
                    : CreateDoorFamilyInstance(
                        doc,
                        projectedPoint,
                        symbolToPlace,
                        hostWall,
                        referenceDirection);
            }
            catch (Exception ex)
            {
                diagnostic = attemptName + ": ошибка вставки: " + ex.Message;
                return false;
            }

            if (createdDoor == null)
            {
                diagnostic = attemptName + ": Revit не создал экземпляр.";
                return false;
            }

            TryRegenerateDoorDocument(doc);

            string roomSideDiagnostic;
            match = GetDoorRoomSideMatch(createdDoor, expectedRoom, out roomSideDiagnostic);

            diagnostic =
                attemptName +
                ", ref = " + FormatVector2D(referenceDirection) +
                ": " + FormatDoorRoomTriplet(createdDoor) +
                " (" + roomSideDiagnostic + ").";

            return true;
        }

        private static bool TryCreateDoorWithAuxiliaryHostWallForRoomSide(Document doc, PreparedDoorPlacement preparedDoor,
            FamilySymbol symbolToPlace, Wall hostWall, XYZ projectedPoint, Level baseLevel, Room expectedRoom,
            List<ElementId> createdElementIds, ApartmentWorksetTargets worksetTargets, out FamilyInstance createdDoor, out string diagnostic)
        {
            createdDoor = null;
            diagnostic = "<нет данных>";

            if (doc == null || preparedDoor == null || symbolToPlace == null || hostWall == null || projectedPoint == null || baseLevel == null)
            {
                diagnostic = "fallback короткая стена-хост: нет документа, двери, стены, точки или уровня.";
                return false;
            }

            WallType auxiliaryWallType = null;
            try
            {
                auxiliaryWallType = hostWall.WallType;
            }
            catch
            {
                auxiliaryWallType = null;
            }

            if (auxiliaryWallType == null)
            {
                diagnostic = "fallback короткая стена-хост: не удалось получить тип исходной стены.";
                return false;
            }

            XYZ wallDirection = GetWallAxisDirection2D(hostWall);
            if (wallDirection == null)
            {
                diagnostic = "fallback короткая стена-хост: не удалось определить ось исходной стены.";
                return false;
            }

            double doorWidth = preparedDoor.DoorWidthMm > 0
                ? IDHelper.ConvertMmToInternal(preparedDoor.DoorWidthMm)
                : IDHelper.ConvertMmToInternal(900);

            double auxiliaryWallLength = Math.Max(doorWidth + IDHelper.ConvertMmToInternal(150), IDHelper.ConvertMmToInternal(1050));
            double wallHeight = GetWallHeightInternal(hostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);

            double z0;
            double z1;
            GetWallVerticalExtents(hostWall, baseLevel, wallHeight, out z0, out z1);
            double wallBaseOffset = GetWallBaseOffsetForLevel(hostWall, baseLevel, z0);

            Opening slotOpening = null;
            List<string> attempts = new List<string>();

            try
            {
                slotOpening = CreateDoorAuxiliaryHostSlotOpening(doc, hostWall, projectedPoint, auxiliaryWallLength, z0, z1);
            }
            catch (Exception ex)
            {
                diagnostic = "fallback слот в основной стене: " + ex.Message;
                return false;
            }

            if (slotOpening == null)
            {
                diagnostic = "fallback слот в основной стене: Revit не создал проём.";
                return false;
            }

            XYZ reversedWallDirection = new XYZ(-wallDirection.X, -wallDirection.Y, 0);
            List<Tuple<string, XYZ>> directions = new List<Tuple<string, XYZ>>
            {
                Tuple.Create("ось короткой стены = -ось исходной", reversedWallDirection),
                Tuple.Create("ось короткой стены = ось исходной", wallDirection)
            };

            foreach (Tuple<string, XYZ> direction in directions)
            {
                if (direction == null || direction.Item2 == null)
                    continue;

                Wall auxiliaryWall = null;
                FamilyInstance candidateDoor = null;

                try
                {
                    Line auxiliaryAxis = BuildDoorAuxiliaryHostWallAxis(WithZ(projectedPoint, 0.0), direction.Item2, auxiliaryWallLength);
                    if (auxiliaryAxis == null)
                    {
                        attempts.Add(direction.Item1 + ": не удалось построить ось.");
                        continue;
                    }

                    auxiliaryWall = Wall.Create(doc, auxiliaryAxis, auxiliaryWallType.Id, baseLevel.Id, wallHeight, 0, false, false);
                    ApplyWallPresetParameters(auxiliaryWall, baseLevel, null, wallBaseOffset, wallHeight);
                    TryAssignElementToWorkset(auxiliaryWall, worksetTargets != null ? worksetTargets.WallWorksetId : null);
                    doc.Regenerate();
                    TryJoinGeometry(doc, hostWall, auxiliaryWall);
                }
                catch (Exception ex)
                {
                    attempts.Add(direction.Item1 + ": короткая стена не создана (" + ex.Message + ").");

                    if (auxiliaryWall != null)
                        TryDeleteElement(doc, auxiliaryWall.Id, null);

                    continue;
                }

                DoorRoomSideMatch match;
                string attemptDiagnostic;
                bool created = TryCreateDoorAndReadRoomSide(
                    doc,
                    "fallback короткая стена-хост, " + direction.Item1,
                    symbolToPlace,
                    auxiliaryWall,
                    projectedPoint,
                    baseLevel,
                    expectedRoom,
                    out candidateDoor,
                    out match,
                    out attemptDiagnostic);

                attempts.Add(attemptDiagnostic);

                if (created && match == DoorRoomSideMatch.ToRoom)
                {
                    createdDoor = candidateDoor;
                    SetWallRoomBounding(auxiliaryWall, false);
                    TryRegenerateDoorDocument(doc);

                    if (createdElementIds != null)
                    {
                        createdElementIds.Add(slotOpening.Id);
                        createdElementIds.Add(auxiliaryWall.Id);
                    }

                    diagnostic =
                        "fallback короткая стена-хост сработал: слот ID = " + IDHelper.ElIdValue(slotOpening.Id) +
                        ", стена ID = " + IDHelper.ElIdValue(auxiliaryWall.Id) +
                        ". Попытки: " + string.Join("; ", attempts.ToArray());

                    return true;
                }

                if (candidateDoor != null)
                    TryDeleteElement(doc, candidateDoor.Id, attempts);

                if (auxiliaryWall != null)
                    TryDeleteElement(doc, auxiliaryWall.Id, attempts);

                TryRegenerateDoorDocument(doc);
            }

            if (slotOpening != null)
                TryDeleteElement(doc, slotOpening.Id, attempts);

            diagnostic = "fallback короткая стена-хост не дал ToRoom. Попытки: " + string.Join("; ", attempts.ToArray());
            return false;
        }

        private static bool SetWallRoomBounding(Wall wall, bool roomBounding)
        {
            if (wall == null)
                return false;

            try
            {
                Parameter p = wall.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                if (p == null || p.IsReadOnly || p.StorageType != StorageType.Integer)
                    return false;

                p.Set(roomBounding ? 1 : 0);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Line BuildDoorAuxiliaryHostWallAxis(XYZ projectedPoint, XYZ wallDirection, double wallLength)
        {
            XYZ direction = Normalize2D(wallDirection);
            if (projectedPoint == null || direction == null || wallLength <= 1e-9)
                return null;

            XYZ half = direction * (wallLength / 2.0);
            XYZ p0 = new XYZ(projectedPoint.X - half.X, projectedPoint.Y - half.Y, projectedPoint.Z);
            XYZ p1 = new XYZ(projectedPoint.X + half.X, projectedPoint.Y + half.Y, projectedPoint.Z);

            return Line.CreateBound(p0, p1);
        }

        private static Opening CreateDoorAuxiliaryHostSlotOpening(Document doc, Wall hostWall, XYZ projectedPoint,
            double openingWidth, double z0, double z1)
        {
            if (doc == null || hostWall == null || projectedPoint == null || openingWidth <= 1e-9 || z1 <= z0)
                return null;

            XYZ wallDirection = GetWallAxisDirection2D(hostWall);
            if (wallDirection == null)
                return null;

            XYZ half = wallDirection * (openingWidth / 2.0);
            XYZ p0 = new XYZ(projectedPoint.X - half.X, projectedPoint.Y - half.Y, z0);
            XYZ p1 = new XYZ(projectedPoint.X + half.X, projectedPoint.Y + half.Y, z1);

            return doc.Create.NewOpening(hostWall, p0, p1);
        }

        private static void GetWallVerticalExtents(Wall wall, Level baseLevel, double wallHeight, out double z0, out double z1)
        {
            z0 = baseLevel != null ? baseLevel.Elevation : 0;
            try
            {
                Parameter pBaseOffset = wall != null ? wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET) : null;
                if (pBaseOffset != null && pBaseOffset.StorageType == StorageType.Double)
                    z0 += pBaseOffset.AsDouble();
            }
            catch
            {
            }

            z1 = z0 + wallHeight;

            try
            {
                BoundingBoxXYZ bbox = wall != null ? wall.get_BoundingBox(null) : null;
                if (bbox != null && bbox.Max.Z > bbox.Min.Z + 1e-6)
                {
                    z0 = bbox.Min.Z;
                    z1 = bbox.Max.Z;
                }
            }
            catch
            {
            }
        }

        private static XYZ AlignHostedFamilyInsertionPointToHostWallBase(XYZ point, Wall hostWall, Level baseLevel)
        {
            if (point == null || hostWall == null)
                return point;

            double wallHeight = GetWallHeightInternal(hostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);

            double z0;
            double z1;
            GetWallVerticalExtents(hostWall, baseLevel, wallHeight, out z0, out z1);

            return WithZ(point, z0);
        }

        private static bool ShouldCreateHostedFamilyWithLevelHostOverload(Wall hostWall, Level baseLevel)
        {
            if (hostWall == null || baseLevel == null)
                return false;

            double wallHeight = GetWallHeightInternal(hostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);

            double z0;
            double z1;
            GetWallVerticalExtents(hostWall, baseLevel, wallHeight, out z0, out z1);

            double tolerance = IDHelper.ConvertMmToInternal(1);
            return Math.Abs(baseLevel.Elevation) > tolerance ||
                   Math.Abs(z0) > tolerance ||
                   Math.Abs(z0 - baseLevel.Elevation) > tolerance;
        }

        private static XYZ GetDoorHostLevelInsertionPoint(XYZ point, Wall hostWall, Level baseLevel)
        {
            if (point == null)
                return null;

            if (baseLevel == null || hostWall == null)
                return point;

            double wallHeight = GetWallHeightInternal(hostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);

            double z0;
            double z1;
            GetWallVerticalExtents(hostWall, baseLevel, wallHeight, out z0, out z1);

            return WithZ(point, z0);
        }

        private static double GetWallBaseOffsetForLevel(Wall wall, Level baseLevel, double wallBaseZ)
        {
            try
            {
                Parameter p = wall != null ? wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET) : null;
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }
            catch
            {
            }

            return baseLevel != null
                ? wallBaseZ - baseLevel.Elevation
                : 0.0;
        }

        private static DoorRoomSideMatch GetDoorRoomSideMatch(FamilyInstance createdDoor, Room expectedRoom,
            out string diagnostic)
        {
            diagnostic = "<нет данных>";

            if (createdDoor == null)
            {
                diagnostic = "нет созданной двери";
                return DoorRoomSideMatch.None;
            }

            Room fromRoom = null;
            Room toRoom = null;

            try
            {
                fromRoom = createdDoor.FromRoom;
            }
            catch (Exception ex)
            {
                diagnostic = "FromRoom созданной двери не прочитался: " + ex.Message;
                return DoorRoomSideMatch.None;
            }

            try
            {
                toRoom = createdDoor.ToRoom;
            }
            catch (Exception ex)
            {
                diagnostic = "ToRoom созданной двери не прочитался: " + ex.Message;
                return DoorRoomSideMatch.None;
            }

            if (expectedRoom == null)
            {
                if (toRoom == null)
                {
                    diagnostic = "совпал ToRoom = <нет>";
                    return DoorRoomSideMatch.ToRoom;
                }

                if (fromRoom == null)
                {
                    diagnostic = "совпал FromRoom = <нет>";
                    return DoorRoomSideMatch.FromRoom;
                }

                diagnostic = "ожидалась сторона без помещения";
                return DoorRoomSideMatch.None;
            }

            long expectedRoomId = IDHelper.ElIdValue(expectedRoom.Id);

            if (RoomIdEquals(toRoom, expectedRoomId))
            {
                diagnostic = "совпал ToRoom";
                return DoorRoomSideMatch.ToRoom;
            }

            if (RoomIdEquals(fromRoom, expectedRoomId))
            {
                diagnostic = "совпал FromRoom";
                return DoorRoomSideMatch.FromRoom;
            }

            diagnostic = "ожидался Room ID = " + expectedRoomId + ", совпадений в FromRoom/ToRoom нет";
            return DoorRoomSideMatch.None;
        }

        private static bool RoomIdEquals(Room room, long expectedRoomId)
        {
            if (room == null)
                return false;

            return IDHelper.ElIdValue(room.Id) == expectedRoomId;
        }

        private static string FormatDoorRoomTriplet(FamilyInstance door)
        {
            Room room = null;
            Room fromRoom = null;
            Room toRoom = null;

            if (door != null)
            {
                try
                {
                    room = door.Room;
                }
                catch
                {
                    room = null;
                }

                try
                {
                    fromRoom = door.FromRoom;
                }
                catch
                {
                    fromRoom = null;
                }

                try
                {
                    toRoom = door.ToRoom;
                }
                catch
                {
                    toRoom = null;
                }
            }

            return "Room = '" + FormatDiagnosticValue(FormatRoomLookupValue(room)) +
                   "', FromRoom = '" + FormatDiagnosticValue(FormatRoomLookupValue(fromRoom)) +
                   "', ToRoom = '" + FormatDiagnosticValue(FormatRoomLookupValue(toRoom)) + "'";
        }

        private bool PlaceEntranceDoorForApartment(Document doc, PreparedDoorPlacement preparedDoor, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, Level baseLevel, List<string> debugMessages, ApartmentProcessState state,
            List<ElementId> createdDoorIds = null, ApartmentWorksetTargets worksetTargets = null)
        {
            if (doc == null || preparedDoor == null || preparedDoor.DoorSymbol == null || preparedDoor.InsertPoint == null || baseLevel == null)
                return false;

            double maxDistanceToWallAxis = IDHelper.ConvertMmToInternal(1000);

            Wall hostWall;
            XYZ projectedPoint;
            double distanceToWallAxis;
            bool hostFromExistingWall;

            bool foundHost = TryFindEntranceDoorHostWall(
                doc,
                preparedDoor,
                createdWallsForApartment,
                existingWallsOnLevel,
                maxDistanceToWallAxis,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis,
                out hostFromExistingWall);

            if (!foundHost || hostWall == null || projectedPoint == null)
            {
                AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Не найдена стена-хост для входной двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                    " квартиры ID = " + IDHelper.ElIdValue(preparedDoor.ApartmentId) + ". " +
                    BuildDoorHostSearchDiagnostic(doc, preparedDoor, existingWallsOnLevel, createdWallsForApartment, maxDistanceToWallAxis));

                return false;
            }

            projectedPoint = AlignHostedFamilyInsertionPointToHostWallBase(projectedPoint, hostWall, baseLevel);

            FamilyInstance apartmentFi = doc.GetElement(preparedDoor.ApartmentId) as FamilyInstance;
            XYZ interiorDirection;
            XYZ outwardDirection;
            string exteriorSideSource;

            if (!TryResolveEntranceDoorExteriorSide(
                apartmentFi,
                doc,
                hostWall,
                projectedPoint,
                preparedDoor.InteriorReferencePoint,
                out exteriorSideSource,
                out interiorDirection,
                out outwardDirection))
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Не удалось определить внешнюю сторону квартиры для входной двери ID = " +
                    FormatElementIdForDiagnostic(preparedDoor.Door2DId) + ".");

                return false;
            }

            double referenceOffset = Math.Max(GetWallHalfWidth(hostWall) + IDHelper.ConvertMmToInternal(300), IDHelper.ConvertMmToInternal(500));
            preparedDoor.InteriorReferencePoint = projectedPoint + interiorDirection * referenceOffset;
            List<string> entranceDiagnostics = new List<string>();

            bool hostAxisReversed = EnsureEntranceHostWallNormalOutside(
                doc,
                hostWall,
                outwardDirection,
                preparedDoor,
                !hostFromExistingWall,
                entranceDiagnostics,
                null);

            FamilySymbol entranceDoorSymbol = preparedDoor.DoorSymbol;
            if (entranceDoorSymbol == null)
            {
                AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Для входной двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                    " не удалось определить итоговый тип перед вставкой.");

                return false;
            }

            string entranceInsertionSideSource;
            XYZ entranceInsertionPoint;
            XYZ entranceReferenceDirection = null;
            bool useEntranceReferenceDirection = false;

            if (hostFromExistingWall)
            {
                entranceInsertionPoint = projectedPoint;
                entranceInsertionSideSource = "ось существующей стены";

                XYZ wallDirection = GetWallAxisDirection2D(hostWall);
                XYZ wallNormal = GetWallAxisNormal2D(hostWall);
                XYZ outward = Normalize2D(outwardDirection);

                if (wallDirection != null && wallNormal != null && outward != null && Dot2D(wallNormal, outward) < -0.25)
                {
                    entranceReferenceDirection = new XYZ(-wallDirection.X, -wallDirection.Y, 0);
                    useEntranceReferenceDirection = true;
                    entranceInsertionSideSource += ", ref=-wallDir";

                    AddApartmentDiagnostic(
                        null,
                        entranceDiagnostics,
                        "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        ": существующая стена имеет нормаль внутрь квартиры, сначала пробую вставку по оси стены с referenceDirection = -wallDir.");
                }
            }
            else
            {
                entranceInsertionPoint = GetEntranceDoorInsertionPoint(
                    apartmentFi,
                    doc,
                    projectedPoint,
                    hostWall,
                    outwardDirection,
                    false,
                    out entranceInsertionSideSource);
            }

            if (entranceInsertionPoint == null)
                entranceInsertionPoint = projectedPoint;

            preparedDoor.DoorSymbol = entranceDoorSymbol;
            preparedDoor.SelectedDoorTypeName = BuildDoorTypeDisplayName(entranceDoorSymbol);

            AddApartmentDiagnostic(
                null,
                entranceDiagnostics,
                "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                ": хост = " + (hostFromExistingWall ? "существующая стена" : "созданная стена") +
                " ID = " + IDHelper.ElIdValue(hostWall.Id) +
                ", расстояние до оси = " + FormatLengthMm(distanceToWallAxis) +
                ", точка вставки = " + FormatPointMm(projectedPoint) +
                ", точка создания = " + FormatPointMm(entranceInsertionPoint) +
                ", сторона создания = " + entranceInsertionSideSource +
                ", referenceDirection = " + (useEntranceReferenceDirection ? FormatVector2D(entranceReferenceDirection) : "<стандартная>") +
                ", внутрь квартиры = " + FormatVector2D(interiorDirection) +
                ", наружу квартиры = " + FormatVector2D(outwardDirection) +
                ", сторона определена = " + exteriorSideSource +
                ", нормаль оси стены = " + FormatVector2D(GetWallAxisNormal2D(hostWall)) +
                ", разворот оси хоста отключён = " + (!hostAxisReversed) +
                ", тип = '" + BuildDoorTypeDisplayName(entranceDoorSymbol) + "'.");

            FamilyInstance createdDoor = null;
            string entrancePlacementFailureDiagnostic = null;

            try
            {
                if (hostFromExistingWall)
                {
                    createdDoor = CreateMatchingEntranceDoorForExistingHost(
                        doc,
                        preparedDoor,
                        hostWall,
                        baseLevel,
                        projectedPoint,
                        outwardDirection,
                        entranceDoorSymbol,
                        entranceDoorSymbol,
                        entranceReferenceDirection,
                        useEntranceReferenceDirection,
                        createdWallsForApartment,
                        createdDoorIds,
                        worksetTargets,
                        entranceDiagnostics,
                        out entrancePlacementFailureDiagnostic);
                }
                else
                {
                    createdDoor = CreateEntranceDoorFamilyInstance(
                        doc,
                        entranceInsertionPoint,
                        projectedPoint,
                        entranceDoorSymbol,
                        hostWall,
                        baseLevel,
                        preparedDoor,
                        outwardDirection,
                        entranceReferenceDirection,
                        useEntranceReferenceDirection,
                        null,
                        entranceDiagnostics,
                        out entrancePlacementFailureDiagnostic);
                }
            }
            catch (Exception ex)
            {
                AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);
                AddApartmentDiagnostics(state, debugMessages, entranceDiagnostics);
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Ошибка вставки входной двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                    " в стену ID = " + IDHelper.ElIdValue(hostWall.Id) +
                    ", тип '" + BuildDoorTypeDisplayName(entranceDoorSymbol) + "': " + ex.Message);

                return false;
            }

            if (createdDoor == null)
            {
                AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);
                AddApartmentDiagnostics(state, debugMessages, entranceDiagnostics);
                AddApartmentDiagnostic(state, debugMessages, entrancePlacementFailureDiagnostic);
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) + " не поставлена: не найден подходящий вариант с направлением наружу" +
                    " в стене ID = " + IDHelper.ElIdValue(hostWall.Id) +
                    ", тип '" + BuildDoorTypeDisplayName(entranceDoorSymbol) + "'.");

                return false;
            }

            if (createdDoor.Symbol != null)
            {
                preparedDoor.DoorSymbol = createdDoor.Symbol;
                preparedDoor.SelectedDoorTypeName = BuildDoorTypeDisplayName(createdDoor.Symbol);
            }

            if (!hostFromExistingWall)
            {
                OrientEntranceDoorBy2DSource(doc, createdDoor, preparedDoor, hostWall, projectedPoint, entranceDiagnostics, null);
                EnsureEntranceDoorFacesOutward(doc, createdDoor, outwardDirection, entranceDiagnostics);
                ApplyDoorPlacementSideDirection(doc, createdDoor, preparedDoor, hostWall, projectedPoint, entranceDiagnostics);

                string entranceValidationDiagnostic;
                if (!IsEntranceDoorPlacementAcceptable(createdDoor, hostWall, outwardDirection, preparedDoor, out entranceValidationDiagnostic))
                {
                    TryDeleteElement(doc, createdDoor.Id, entranceDiagnostics);
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);
                    AddApartmentDiagnostics(state, debugMessages, entranceDiagnostics);
                    AddApartmentDiagnostic(state, debugMessages, entranceValidationDiagnostic);
                    return false;
                }
            }

            if (state != null)
                state.InstalledEntranceDoorsCount++;

            TryAssignElementToWorkset(createdDoor, worksetTargets != null ? worksetTargets.DoorWorksetId : null);

            if (createdDoorIds != null)
                createdDoorIds.Add(createdDoor.Id);

            return true;
        }

        private static void AddPreparedDoorDiagnostics(ApartmentProcessState state, List<string> debugMessages, PreparedDoorPlacement preparedDoor)
        {
            if (preparedDoor == null || preparedDoor.Diagnostics == null || preparedDoor.Diagnostics.Count == 0)
                return;

            foreach (string diagnostic in preparedDoor.Diagnostics)
                AddApartmentDiagnostic(state, debugMessages, diagnostic);
        }

        private string BuildDoorHostSearchDiagnostic(Document doc, PreparedDoorPlacement preparedDoor, List<ExistingWallLineInfo> existingWallsOnLevel,
            List<Wall> createdWallsForApartment, double maxDistanceToWallAxis)
        {
            if (preparedDoor == null)
                return "";

            List<Wall> existingWallCandidates = GetExistingWallsFromLineInfo(doc, existingWallsOnLevel);

            string existingText = BuildNearestWallDiagnostic(
                preparedDoor.InsertPoint,
                existingWallCandidates,
                true,
                maxDistanceToWallAxis);

            string createdText = BuildNearestWallDiagnostic(
                preparedDoor.InsertPoint,
                createdWallsForApartment,
                false,
                maxDistanceToWallAxis);

            return "Точка 2D = " + FormatPointMm(preparedDoor.InsertPoint) +
                   ", существующих стен-кандидатов = " + existingWallCandidates.Count +
                   ", созданных стен-кандидатов = " + (createdWallsForApartment != null ? createdWallsForApartment.Count : 0) +
                   ". Ближайшая существующая: " + existingText +
                   ". Ближайшая созданная: " + createdText;
        }

        private string BuildNearestWallDiagnostic(XYZ doorPoint, List<Wall> walls, bool includeWallHalfWidth, double maxDistanceToWallAxis)
        {
            if (doorPoint == null)
                return "нет точки двери";

            if (walls == null || walls.Count == 0)
                return "нет стен";

            Wall nearestWall = null;
            XYZ nearestProjectedPoint = null;
            double nearestDistance = double.MaxValue;

            foreach (Wall wall in walls)
            {
                if (wall == null)
                    continue;

                Line wallLine = GetWallAxisLine(wall);
                if (wallLine == null)
                    continue;

                XYZ projectedPoint;
                double distance;
                if (!TryProjectPointToSegment2D(doorPoint, wallLine, out projectedPoint, out distance))
                    continue;

                if (distance < nearestDistance)
                {
                    nearestDistance = distance;
                    nearestWall = wall;
                    nearestProjectedPoint = projectedPoint;
                }
            }

            if (nearestWall == null)
                return "не удалось спроецировать на стены";

            double wallHalfWidth = includeWallHalfWidth ? GetWallHalfWidth(nearestWall) : 0;
            double allowedDistance = maxDistanceToWallAxis + wallHalfWidth;

            return "ID = " + IDHelper.ElIdValue(nearestWall.Id) +
                   ", расстояние до оси = " + FormatLengthMm(nearestDistance) +
                   ", половина толщины = " + FormatLengthMm(wallHalfWidth) +
                   ", допуск = " + FormatLengthMm(allowedDistance) +
                   ", проекция = " + FormatPointMm(nearestProjectedPoint);
        }

        private static FamilyInstance CreateDoorFamilyInstance(Document doc, XYZ insertionPoint, FamilySymbol symbolToPlace, Wall hostWall,
            XYZ referenceDirection)
        {
            if (doc == null || insertionPoint == null || symbolToPlace == null || hostWall == null || referenceDirection == null)
                return null;

            return doc.Create.NewFamilyInstance(
                insertionPoint,
                symbolToPlace,
                referenceDirection,
                hostWall,
                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }

        private static XYZ GetEntranceDoorInsertionPoint(FamilyInstance apartmentFi, Document doc, XYZ projectedPoint, Wall hostWall,
            XYZ outwardDirection, bool hostFromExistingWall, out string sideSource)
        {
            sideSource = "<нет>";

            if (projectedPoint == null)
                return null;

            if (!hostFromExistingWall)
            {
                sideSource = "ось созданной стены";
                return projectedPoint;
            }

            XYZ outsideDirection = ResolveNoApartmentRoomSideDirection(apartmentFi, doc, projectedPoint, hostWall, outwardDirection, out sideSource);
            if (outsideDirection == null)
            {
                sideSource = "ось существующей стены";
                return projectedPoint;
            }

            double halfWidth = GetWallHalfWidth(hostWall);
            if (halfWidth <= 0)
            {
                sideSource = sideSource + ", без толщины стены";
                return projectedPoint;
            }

            double offset = halfWidth + IDHelper.ConvertMmToInternal(5);
            if (offset <= 1e-9)
                return projectedPoint;

            return projectedPoint + outsideDirection * offset;
        }

        private static XYZ ResolveNoApartmentRoomSideDirection(FamilyInstance apartmentFi, Document doc, XYZ projectedPoint, Wall hostWall,
            XYZ outwardDirection, out string sideSource)
        {
            sideSource = "<нет>";

            if (projectedPoint == null || hostWall == null)
                return null;

            XYZ wallNormal = GetWallAxisNormal2D(hostWall);
            if (wallNormal == null)
                return Normalize2D(outwardDirection);

            XYZ oppositeNormal = new XYZ(-wallNormal.X, -wallNormal.Y, 0);
            double sampleOffset = Math.Max(GetWallHalfWidth(hostWall) + IDHelper.ConvertMmToInternal(300), IDHelper.ConvertMmToInternal(500));

            bool normalInside = IsPointInsideApartment2DRooms(apartmentFi, doc, projectedPoint + wallNormal * sampleOffset);
            bool oppositeInside = IsPointInsideApartment2DRooms(apartmentFi, doc, projectedPoint + oppositeNormal * sampleOffset);

            if (normalInside != oppositeInside)
            {
                sideSource = normalInside
                    ? "сторона без помещений квартиры (-нормаль стены)"
                    : "сторона без помещений квартиры (+нормаль стены)";

                return normalInside ? oppositeNormal : wallNormal;
            }

            XYZ outward = Normalize2D(outwardDirection);
            if (outward != null)
            {
                sideSource = "наружу квартиры по fallback";
                return outward;
            }

            return null;
        }

        private static bool EnsureEntranceHostWallNormalOutside(Document doc, Wall hostWall, XYZ outwardDirection, PreparedDoorPlacement preparedDoor,
            bool allowReverseHostAxis, List<string> debugMessages, ApartmentProcessState state)
        {
            if (hostWall == null || outwardDirection == null)
                return false;

            XYZ outward = Normalize2D(outwardDirection);
            XYZ wallNormal = GetWallAxisNormal2D(hostWall);

            if (outward == null || wallNormal == null)
                return false;

            if (Dot2D(wallNormal, outward) > 0.25)
                return false;

            AddApartmentDiagnostic(
                state,
                debugMessages,
                "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor != null ? preparedDoor.Door2DId : ElementId.InvalidElementId) +
                ": разворот оси стены-хоста отключён. Нормаль стены = " +
                FormatVector2D(wallNormal) + ", наружу квартиры = " + FormatVector2D(outward) + ".");

            return false;
        }

        private FamilyInstance CreateEntranceDoorFamilyInstance(Document doc, XYZ insertionPoint, XYZ projectedPoint, FamilySymbol baseSymbol, Wall hostWall,
            Level baseLevel, PreparedDoorPlacement preparedDoor, XYZ outwardDirection, XYZ referenceDirection, bool useReferenceDirection,
            ApartmentProcessState state, List<string> debugMessages, out string failureDiagnostic)
        {
            failureDiagnostic = null;

            if (doc == null || insertionPoint == null || projectedPoint == null || baseSymbol == null || hostWall == null || baseLevel == null)
                return null;

            try
            {
                if (!baseSymbol.IsActive)
                {
                    baseSymbol.Activate();
                    doc.Regenerate();
                }
            }
            catch (Exception ex)
            {
                failureDiagnostic =
                    "Не удалось активировать тип входной двери '" + BuildDoorTypeDisplayName(baseSymbol) + "': " + ex.Message;
                return null;
            }

            FamilyInstance createdDoor = null;
            bool useLevelHostOverload = ShouldCreateHostedFamilyWithLevelHostOverload(hostWall, baseLevel);
            if (useReferenceDirection && !useLevelHostOverload)
            {
                createdDoor = TryCreateDoorWithReferenceDirection(doc, insertionPoint, baseSymbol, referenceDirection, hostWall, baseLevel);
                if (createdDoor == null)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor != null ? preparedDoor.Door2DId : ElementId.InvalidElementId) +
                        ": вставка с referenceDirection = " + FormatVector2D(referenceDirection) +
                        " не создалась, пробую стандартную вставку.");
                }
            }

            if (createdDoor == null)
                createdDoor = TryCreateDoorDefault(doc, insertionPoint, baseSymbol, hostWall, baseLevel);

            if (createdDoor == null)
            {
                failureDiagnostic =
                    "Revit не создал входную дверь в точке " + FormatPointMm(insertionPoint) +
                    ", тип '" + BuildDoorTypeDisplayName(baseSymbol) + "'.";
                return null;
            }

            OrientEntranceDoorBy2DSource(doc, createdDoor, preparedDoor, hostWall, projectedPoint, null, null);
            EnsureEntranceDoorFacesOutward(doc, createdDoor, outwardDirection, null);
            ApplyDoorPlacementSideDirection(doc, createdDoor, preparedDoor, hostWall, projectedPoint, debugMessages);

            return createdDoor;
        }

        private class EntranceDoorCreationCandidate
        {
            public FamilySymbol Symbol { get; set; }
            public XYZ ReferenceDirection { get; set; }
            public bool UseReferenceDirection { get; set; }
            public string Name { get; set; }
        }

        private FamilyInstance CreateMatchingEntranceDoorForExistingHost(Document doc, PreparedDoorPlacement preparedDoor, Wall hostWall, Level baseLevel,
            XYZ projectedPoint, XYZ outwardDirection, FamilySymbol sourceSymbol, FamilySymbol preferredSymbol, XYZ firstReferenceDirection,
            bool useFirstReferenceDirection, List<Wall> createdWallsForApartment, List<ElementId> createdElementIds,
            ApartmentWorksetTargets worksetTargets, List<string> diagnostics, out string failureDiagnostic)
        {
            failureDiagnostic = null;

            if (doc == null || preparedDoor == null || hostWall == null || baseLevel == null || projectedPoint == null)
                return null;

            List<EntranceDoorCreationCandidate> candidates = BuildEntranceDoorCreationCandidates(
                doc,
                hostWall,
                sourceSymbol,
                preferredSymbol,
                firstReferenceDirection,
                useFirstReferenceDirection);

            List<string> attempts = new List<string>();

            foreach (EntranceDoorCreationCandidate candidate in candidates)
            {
                if (candidate == null || candidate.Symbol == null)
                    continue;

                string candidateFailure;
                List<string> candidateDiagnostics = new List<string>();
                FamilyInstance candidateDoor = CreateEntranceDoorFamilyInstance(
                    doc,
                    projectedPoint,
                    projectedPoint,
                    candidate.Symbol,
                    hostWall,
                    baseLevel,
                    preparedDoor,
                    outwardDirection,
                    candidate.ReferenceDirection,
                    candidate.UseReferenceDirection,
                    null,
                    candidateDiagnostics,
                    out candidateFailure);

                if (candidateDoor == null)
                {
                    attempts.Add(candidate.Name + ": не создано" + (!string.IsNullOrWhiteSpace(candidateFailure) ? " (" + candidateFailure + ")" : ""));
                    AddAttemptDiagnostics(attempts, candidateDiagnostics, candidate.Name);
                    continue;
                }

                string validationDiagnostic;
                if (IsEntranceDoorPlacementAcceptable(candidateDoor, hostWall, outwardDirection, preparedDoor, out validationDiagnostic))
                    return candidateDoor;

                attempts.Add(candidate.Name + ": " + validationDiagnostic);
                AddAttemptDiagnostics(attempts, candidateDiagnostics, candidate.Name);
                TryDeleteElement(doc, candidateDoor.Id, null);
            }

            FamilyInstance auxiliaryDoor = TryCreateEntranceDoorWithAuxiliaryHostWall(
                doc,
                preparedDoor,
                hostWall,
                baseLevel,
                projectedPoint,
                outwardDirection,
                createdWallsForApartment,
                createdElementIds,
                candidates,
                worksetTargets,
                attempts);

            if (auxiliaryDoor != null)
                return auxiliaryDoor;

            failureDiagnostic =
                "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                " не поставлена в существующую стену: ни один вариант тип/referenceDirection и fallback через короткую стену-хост не дал нужную сторону и разворот. Попытки: " +
                string.Join("; ", attempts.Take(12).ToArray()) +
                (attempts.Count > 12 ? "; ..." : "") + ".";

            return null;
        }

        private FamilyInstance TryCreateEntranceDoorWithAuxiliaryHostWall(Document doc, PreparedDoorPlacement preparedDoor, Wall existingHostWall,
            Level baseLevel, XYZ projectedPoint, XYZ outwardDirection, List<Wall> createdWallsForApartment, List<ElementId> createdElementIds,
            List<EntranceDoorCreationCandidate> candidates, ApartmentWorksetTargets worksetTargets, List<string> attempts)
        {
            if (doc == null || preparedDoor == null || existingHostWall == null || baseLevel == null || projectedPoint == null || candidates == null)
                return null;

            WallType auxiliaryWallType = ResolveEntranceAuxiliaryHostWallType(existingHostWall, createdWallsForApartment);
            if (auxiliaryWallType == null)
            {
                if (attempts != null)
                    attempts.Add("fallback короткая стена-хост: не найден тип стены");
                return null;
            }

            double doorWidth = preparedDoor.DoorWidthMm > 0
                ? IDHelper.ConvertMmToInternal(preparedDoor.DoorWidthMm)
                : IDHelper.ConvertMmToInternal(900);
            double hostWallLength = Math.Max(doorWidth + IDHelper.ConvertMmToInternal(150), IDHelper.ConvertMmToInternal(1050));
            double openingWidth = hostWallLength;

            Line auxiliaryAxis = BuildEntranceAuxiliaryHostWallAxis(WithZ(projectedPoint, 0.0), existingHostWall, outwardDirection, hostWallLength);
            if (auxiliaryAxis == null)
            {
                if (attempts != null)
                    attempts.Add("fallback короткая стена-хост: не удалось построить ось");
                return null;
            }

            double wallHeight = GetWallHeightInternal(existingHostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);
            double z0;
            double z1;
            GetWallVerticalExtents(existingHostWall, baseLevel, wallHeight, out z0, out z1);
            double wallBaseOffset = GetWallBaseOffsetForLevel(existingHostWall, baseLevel, z0);

            Opening opening = null;
            Wall auxiliaryWall = null;

            try
            {
                opening = CreateEntranceOpeningInExistingWall(doc, existingHostWall, projectedPoint, baseLevel, openingWidth, candidates.FirstOrDefault(x => x != null && x.Symbol != null)?.Symbol);
            }
            catch (Exception ex)
            {
                if (attempts != null)
                    attempts.Add("fallback проём в существующей стене: " + ex.Message);
                opening = null;
            }

            if (opening == null)
            {
                if (attempts != null)
                    attempts.Add("fallback проём в существующей стене: Revit не создал проём");
                return null;
            }

            try
            {
                auxiliaryWall = Wall.Create(doc, auxiliaryAxis, auxiliaryWallType.Id, baseLevel.Id, wallHeight, 0, false, false);
                ApplyWallPresetParameters(auxiliaryWall, baseLevel, null, wallBaseOffset, wallHeight);
                TryAssignElementToWorkset(auxiliaryWall, worksetTargets != null ? worksetTargets.WallWorksetId : null);
                doc.Regenerate();
                TryJoinGeometry(doc, existingHostWall, auxiliaryWall);
            }
            catch (Exception ex)
            {
                if (attempts != null)
                    attempts.Add("fallback короткая стена-хост: не создана (" + ex.Message + ")");

                if (opening != null)
                    TryDeleteElement(doc, opening.Id, null);

                return null;
            }

            foreach (EntranceDoorCreationCandidate candidate in candidates)
            {
                if (candidate == null || candidate.Symbol == null)
                    continue;

                string candidateFailure;
                List<string> candidateDiagnostics = new List<string>();
                FamilyInstance candidateDoor = CreateEntranceDoorFamilyInstance(
                    doc,
                    projectedPoint,
                    projectedPoint,
                    candidate.Symbol,
                    auxiliaryWall,
                    baseLevel,
                    preparedDoor,
                    outwardDirection,
                    candidate.ReferenceDirection,
                    candidate.UseReferenceDirection,
                    null,
                    candidateDiagnostics,
                    out candidateFailure);

                string candidateName = "fallback короткая стена-хост, " + candidate.Name;

                if (candidateDoor == null)
                {
                    if (attempts != null)
                        attempts.Add(candidateName + ": не создано" + (!string.IsNullOrWhiteSpace(candidateFailure) ? " (" + candidateFailure + ")" : ""));
                    AddAttemptDiagnostics(attempts, candidateDiagnostics, candidateName);
                    continue;
                }

                string validationDiagnostic;
                if (IsEntranceDoorPlacementAcceptable(candidateDoor, auxiliaryWall, outwardDirection, preparedDoor, out validationDiagnostic))
                {
                    if (createdElementIds != null)
                    {
                        if (opening != null)
                            createdElementIds.Add(opening.Id);
                        createdElementIds.Add(auxiliaryWall.Id);
                    }

                    return candidateDoor;
                }

                if (attempts != null)
                    attempts.Add(candidateName + ": " + validationDiagnostic);
                AddAttemptDiagnostics(attempts, candidateDiagnostics, candidateName);
                TryDeleteElement(doc, candidateDoor.Id, null);
            }

            if (opening != null)
                TryDeleteElement(doc, opening.Id, null);
            if (auxiliaryWall != null)
                TryDeleteElement(doc, auxiliaryWall.Id, null);

            return null;
        }

        private List<EntranceDoorCreationCandidate> BuildEntranceDoorCreationCandidates(Document doc, Wall hostWall, FamilySymbol sourceSymbol,
            FamilySymbol preferredSymbol, XYZ firstReferenceDirection, bool useFirstReferenceDirection)
        {
            List<EntranceDoorCreationCandidate> result = new List<EntranceDoorCreationCandidate>();

            List<FamilySymbol> symbols = new List<FamilySymbol>();
            AddUniqueFamilySymbol(symbols, preferredSymbol);
            AddUniqueFamilySymbol(symbols, sourceSymbol);

            List<Tuple<string, XYZ, bool>> references = new List<Tuple<string, XYZ, bool>>();
            if (useFirstReferenceDirection)
                AddUniqueReferenceDirection(references, "ref текущий", firstReferenceDirection, true);

            AddUniqueReferenceDirection(references, "стандартная вставка", null, false);

            XYZ wallDirection = GetWallAxisDirection2D(hostWall);
            AddUniqueReferenceDirection(references, "ref wallDir", wallDirection, true);
            if (wallDirection != null)
                AddUniqueReferenceDirection(references, "ref -wallDir", new XYZ(-wallDirection.X, -wallDirection.Y, 0), true);

            foreach (FamilySymbol symbol in symbols)
            {
                foreach (Tuple<string, XYZ, bool> reference in references)
                {
                    result.Add(new EntranceDoorCreationCandidate
                    {
                        Symbol = symbol,
                        ReferenceDirection = reference.Item2,
                        UseReferenceDirection = reference.Item3,
                        Name = "'" + BuildDoorTypeDisplayName(symbol) + "', " + reference.Item1
                    });
                }
            }

            return result;
        }

        private static WallType ResolveEntranceAuxiliaryHostWallType(Wall existingHostWall, List<Wall> createdWallsForApartment)
        {
            try
            {
                WallType existingWallType = existingHostWall != null ? existingHostWall.WallType : null;
                if (existingWallType != null)
                    return existingWallType;
            }
            catch
            {
            }

            if (createdWallsForApartment != null)
            {
                foreach (Wall wall in createdWallsForApartment)
                {
                    if (wall == null)
                        continue;

                    try
                    {
                        if (wall.WallType != null)
                            return wall.WallType;
                    }
                    catch
                    {
                    }
                }
            }

            return null;
        }

        private static Line BuildEntranceAuxiliaryHostWallAxis(XYZ projectedPoint, Wall existingHostWall, XYZ outwardDirection, double wallLength)
        {
            if (projectedPoint == null || existingHostWall == null || outwardDirection == null || wallLength <= 1e-9)
                return null;

            XYZ wallDir = GetWallAxisDirection2D(existingHostWall);
            XYZ outward = Normalize2D(outwardDirection);
            if (wallDir == null || outward == null)
                return null;

            XYZ normal = new XYZ(-wallDir.Y, wallDir.X, 0);
            if (Dot2D(normal, outward) < 0)
                wallDir = new XYZ(-wallDir.X, -wallDir.Y, 0);

            XYZ half = wallDir * (wallLength / 2.0);
            XYZ p0 = new XYZ(projectedPoint.X - half.X, projectedPoint.Y - half.Y, projectedPoint.Z);
            XYZ p1 = new XYZ(projectedPoint.X + half.X, projectedPoint.Y + half.Y, projectedPoint.Z);

            return Line.CreateBound(p0, p1);
        }

        private static Opening CreateEntranceOpeningInExistingWall(Document doc, Wall existingHostWall, XYZ projectedPoint, Level baseLevel,
            double openingWidth, FamilySymbol preferredSymbol)
        {
            if (doc == null || existingHostWall == null || projectedPoint == null || baseLevel == null || openingWidth <= 1e-9)
                return null;

            XYZ wallDir = GetWallAxisDirection2D(existingHostWall);
            if (wallDir == null)
                return null;

            double doorHeight = GetDoorHeightInternal(preferredSymbol);
            if (doorHeight <= IDHelper.ConvertMmToInternal(500))
                doorHeight = IDHelper.ConvertMmToInternal(2200);

            double wallHeight = GetWallHeightInternal(existingHostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = doorHeight + IDHelper.ConvertMmToInternal(300);

            double z0;
            double wallTopZ;
            GetWallVerticalExtents(existingHostWall, baseLevel, wallHeight, out z0, out wallTopZ);
            double z1 = z0 + doorHeight + IDHelper.ConvertMmToInternal(150);
            if (wallTopZ > z0 + IDHelper.ConvertMmToInternal(300))
                z1 = Math.Min(z1, wallTopZ);

            XYZ half = wallDir * (openingWidth / 2.0);

            XYZ p0 = new XYZ(projectedPoint.X - half.X, projectedPoint.Y - half.Y, z0);
            XYZ p1 = new XYZ(projectedPoint.X + half.X, projectedPoint.Y + half.Y, z1);

            return doc.Create.NewOpening(existingHostWall, p0, p1);
        }

        private static double GetDoorHeightInternal(FamilySymbol symbol)
        {
            if (symbol == null)
                return 0;

            double heightInternal;
            if (TryGetLengthParamFromElementOrType(symbol, out heightInternal, "Высота", "Height"))
                return heightInternal;

            try
            {
                Parameter p = symbol.get_Parameter(BuiltInParameter.DOOR_HEIGHT);
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }
            catch
            {
            }

            return 0;
        }

        private static double GetWallHeightInternal(Wall wall)
        {
            if (wall == null)
                return 0;

            try
            {
                Parameter p = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    double value = p.AsDouble();
                    if (value > 1e-9)
                        return value;
                }
            }
            catch
            {
            }

            try
            {
                BoundingBoxXYZ bbox = wall.get_BoundingBox(null);
                if (bbox != null)
                    return Math.Abs(bbox.Max.Z - bbox.Min.Z);
            }
            catch
            {
            }

            return 0;
        }

        private static void AddUniqueFamilySymbol(List<FamilySymbol> symbols, FamilySymbol symbol)
        {
            if (symbols == null || symbol == null)
                return;

            long id = IDHelper.ElIdValue(symbol.Id);
            if (symbols.Any(x => x != null && IDHelper.ElIdValue(x.Id) == id))
                return;

            symbols.Add(symbol);
        }

        private static void AddUniqueReferenceDirection(List<Tuple<string, XYZ, bool>> references, string name, XYZ referenceDirection, bool useReferenceDirection)
        {
            if (references == null)
                return;

            XYZ normalized = useReferenceDirection ? Normalize2D(referenceDirection) : null;
            if (useReferenceDirection && normalized == null)
                return;

            if (!useReferenceDirection)
            {
                if (references.Any(x => x != null && !x.Item3))
                    return;

                references.Add(Tuple.Create(name, (XYZ)null, false));
                return;
            }

            if (references.Any(x => x != null && x.Item3 && Dot2D(x.Item2, normalized) > 0.999))
                return;

            references.Add(Tuple.Create(name, normalized, true));
        }

        private static bool IsEntranceDoorPlacementAcceptable(FamilyInstance createdDoor, Wall hostWall, XYZ outwardDirection,
            PreparedDoorPlacement preparedDoor, out string diagnostic)
        {
            diagnostic = null;

            if (createdDoor == null || outwardDirection == null)
            {
                diagnostic = "Входная дверь не проверена: нет созданного экземпляра или направления наружу.";
                return false;
            }

            XYZ outward = Normalize2D(outwardDirection);
            XYZ wallNormal = GetWallAxisNormal2D(hostWall);
            XYZ facing = GetFamilyInstanceFacingDirection2D(createdDoor, null);
            XYZ hand = GetFamilyInstanceHandDirection2D(createdDoor, null);
            XYZ sourceHand = preparedDoor != null ? Normalize2D(preparedDoor.SourceHandDirection) : null;

            double wallNormalScore = wallNormal != null && outward != null ? Dot2D(wallNormal, outward) : double.NegativeInfinity;
            double facingScore = facing != null && outward != null ? Dot2D(facing, outward) : double.NegativeInfinity;
            double handScore = sourceHand != null && hand != null ? Dot2D(sourceHand, hand) : 1.0;

            bool sideOk = wallNormalScore > 0.25 || facingScore > 0.25;
            bool handOk = sourceHand == null || hand == null || handScore > 0.25;

            if (sideOk && handOk)
                return true;

            diagnostic =
                "не совпало: сторона " + (sideOk ? "OK" : "не OK") +
                " (нормаль стены к наружу = " + FormatDouble(wallNormalScore) +
                ", facing к наружу = " + FormatDouble(facingScore) +
                "), разворот/тип " + (handOk ? "OK" : "не OK") +
                " (hand к 2D = " + FormatDouble(handScore) +
                "), тип = '" + BuildDoorTypeDisplayName(createdDoor.Symbol) +
                "', hand = " + FormatVector2D(hand) +
                ", facing = " + FormatVector2D(facing) + "";

            return false;
        }

        private static void AddAttemptDiagnostics(List<string> attempts, List<string> diagnostics, string candidateName)
        {
            if (attempts == null || diagnostics == null || diagnostics.Count == 0)
                return;

            foreach (string diagnostic in diagnostics)
            {
                if (string.IsNullOrWhiteSpace(diagnostic))
                    continue;

                attempts.Add(candidateName + ": " + diagnostic);
            }
        }

        private static bool TryDeleteElement(Document doc, ElementId elementId, List<string> diagnostics)
        {
            if (doc == null || elementId == null || elementId == ElementId.InvalidElementId)
                return false;

            try
            {
                doc.Delete(elementId);
                return true;
            }
            catch (Exception ex)
            {
                if (diagnostics != null)
                    diagnostics.Add("Не удалось удалить неподходящий вариант двери ID = " + IDHelper.ElIdValue(elementId) + ": " + ex.Message);

                return false;
            }
        }

        private static bool TryJoinGeometry(Document doc, Element first, Element second)
        {
            if (doc == null || first == null || second == null)
                return false;

            try
            {
                if (JoinGeometryUtils.AreElementsJoined(doc, first, second))
                    return true;
            }
            catch
            {
            }

            try
            {
                JoinGeometryUtils.JoinGeometry(doc, first, second);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static FamilyInstance TryCreateDoorWithReferenceDirection(Document doc, XYZ projectedPoint, FamilySymbol symbolToPlace,
            XYZ referenceDirection, Wall hostWall, Level baseLevel)
        {
            if (doc == null || projectedPoint == null || symbolToPlace == null || hostWall == null || referenceDirection == null)
                return null;

            try
            {
                return doc.Create.NewFamilyInstance(
                    projectedPoint,
                    symbolToPlace,
                    referenceDirection,
                    hostWall,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            }
            catch
            {
                return null;
            }
        }

        private static FamilyInstance TryCreateDoorDefault(Document doc, XYZ projectedPoint, FamilySymbol symbolToPlace, Wall hostWall, Level baseLevel)
        {
            if (doc == null || projectedPoint == null || symbolToPlace == null || hostWall == null || baseLevel == null)
                return null;

            XYZ insertionPoint = GetDoorHostLevelInsertionPoint(projectedPoint, hostWall, baseLevel);
            if (insertionPoint == null)
                return null;

            try
            {
                return doc.Create.NewFamilyInstance(
                    insertionPoint,
                    symbolToPlace,
                    hostWall,
                    baseLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            }
            catch
            {
                return null;
            }
        }

        private static double GetDoorFacingOutwardScore(FamilyInstance door, XYZ outwardDirection)
        {
            if (door == null || outwardDirection == null)
                return double.NegativeInfinity;

            XYZ facing = GetFamilyInstanceFacingDirection2D(door, null);
            if (facing == null)
                return double.NegativeInfinity;

            return Dot2D(facing, outwardDirection);
        }

        private static void EnsureEntranceDoorFacesOutward(Document doc, FamilyInstance createdDoor, XYZ outwardDirection, List<string> debugMessages)
        {
            if (createdDoor == null || outwardDirection == null)
                return;

            if (GetDoorFacingOutwardScore(createdDoor, outwardDirection) > 0)
                return;

            if (!CanFlipFacing(createdDoor))
            {
                if (debugMessages != null)
                    debugMessages.Add("Входная дверь ID = " + IDHelper.ElIdValue(createdDoor.Id) + " не поддерживает flipFacing для разворота наружу.");
                return;
            }

            try
            {
                createdDoor.flipFacing();
                if (doc != null)
                    doc.Regenerate();
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось развернуть входную дверь ID = " + IDHelper.ElIdValue(createdDoor.Id) + " наружу: " + ex.Message);
            }
        }

        private bool TryFindEntranceDoorHostWall(Document doc, PreparedDoorPlacement preparedDoor, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, double maxDistanceToWallAxis, out Wall hostWall, out XYZ projectedPoint,
            out double distanceToWallAxis, out bool hostFromExistingWall)
        {
            hostWall = null;
            projectedPoint = null;
            distanceToWallAxis = double.MaxValue;
            hostFromExistingWall = false;

            if (preparedDoor == null || preparedDoor.InsertPoint == null)
                return false;

            List<Wall> existingWallCandidates = GetExistingWallsFromLineInfo(doc, existingWallsOnLevel);
            if (TryFindBestEntranceDoorHostWall(
                preparedDoor,
                existingWallCandidates,
                maxDistanceToWallAxis,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis))
            {
                hostFromExistingWall = true;
                return true;
            }

            if (TryFindBestEntranceDoorHostWall(
                preparedDoor,
                createdWallsForApartment,
                maxDistanceToWallAxis,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis))
            {
                hostFromExistingWall = false;
                return true;
            }

            return false;
        }

        private bool TryFindBestEntranceDoorHostWall(PreparedDoorPlacement preparedDoor, List<Wall> candidateWalls, double maxDistanceToWallAxis,
            out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            bestWall = null;
            bestProjectedPoint = null;
            bestDistance = double.MaxValue;
            double bestScore = double.MaxValue;

            if (preparedDoor == null || preparedDoor.InsertPoint == null || candidateWalls == null || candidateWalls.Count == 0)
                return false;

            foreach (Wall wall in candidateWalls)
            {
                if (wall == null)
                    continue;

                Line wallLine = GetWallAxisLine(wall);
                if (wallLine == null || wallLine.Length < 1e-9)
                    continue;

                XYZ wallDir = Normalize2D(wallLine.GetEndPoint(1) - wallLine.GetEndPoint(0));
                if (wallDir == null)
                    continue;

                XYZ projectedPoint;
                double distance;
                if (!TryProjectPointToSegment2D(preparedDoor.InsertPoint, wallLine, out projectedPoint, out distance))
                    continue;

                double wallHalfWidth = GetWallHalfWidth(wall);
                double allowedDistance = maxDistanceToWallAxis + wallHalfWidth;
                if (distance > allowedDistance)
                    continue;

                double score = Math.Max(0, distance - wallHalfWidth) +
                               GetEntranceDoorHostOrientationPenalty(preparedDoor, wallDir);

                if (score < bestScore || (Math.Abs(score - bestScore) < 1e-9 && distance < bestDistance))
                {
                    bestScore = score;
                    bestDistance = distance;
                    bestWall = wall;
                    bestProjectedPoint = projectedPoint;
                }
            }

            return bestWall != null && bestProjectedPoint != null;
        }

        private static double GetEntranceDoorHostOrientationPenalty(PreparedDoorPlacement preparedDoor, XYZ wallDir)
        {
            if (preparedDoor == null || wallDir == null)
                return 0;

            double bestAlignment = 0;
            bool hasDirection = false;

            XYZ sourceHand = Normalize2D(preparedDoor.SourceHandDirection);
            if (sourceHand != null)
            {
                bestAlignment = Math.Max(bestAlignment, Math.Abs(Dot2D(sourceHand, wallDir)));
                hasDirection = true;
            }

            XYZ sourceFacing = Normalize2D(preparedDoor.SourceFacingDirection);
            if (sourceFacing != null)
            {
                XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0);
                bestAlignment = Math.Max(bestAlignment, Math.Abs(Dot2D(sourceFacing, wallNormal)));
                hasDirection = true;
            }

            if (!hasDirection)
                return 0;

            bestAlignment = Math.Max(0, Math.Min(1, bestAlignment));
            return (1.0 - bestAlignment) * IDHelper.ConvertMmToInternal(300);
        }

        private static bool TryResolveEntranceDoorExteriorSide(FamilyInstance apartmentFi, Document doc, Wall hostWall, XYZ projectedPoint,
            XYZ fallbackInteriorPoint, out string exteriorSideSource, out XYZ interiorDirection, out XYZ outwardDirection)
        {
            exteriorSideSource = null;
            interiorDirection = null;
            outwardDirection = null;

            if (hostWall == null || projectedPoint == null)
                return false;

            XYZ wallNormal = GetWallAxisNormal2D(hostWall);
            if (wallNormal == null)
                return false;

            XYZ oppositeNormal = new XYZ(-wallNormal.X, -wallNormal.Y, 0);
            double sampleOffset = Math.Max(GetWallHalfWidth(hostWall) + IDHelper.ConvertMmToInternal(300), IDHelper.ConvertMmToInternal(500));

            bool normalInside = IsPointInsideApartment2DRooms(apartmentFi, doc, projectedPoint + wallNormal * sampleOffset);
            bool oppositeInside = IsPointInsideApartment2DRooms(apartmentFi, doc, projectedPoint + oppositeNormal * sampleOffset);

            if (normalInside != oppositeInside)
            {
                interiorDirection = normalInside ? wallNormal : oppositeNormal;
                outwardDirection = normalInside ? oppositeNormal : wallNormal;
                exteriorSideSource = "контур помещений квартиры";
                return true;
            }

            XYZ fallbackInteriorDirection = null;
            if (fallbackInteriorPoint != null)
                fallbackInteriorDirection = Normalize2D(fallbackInteriorPoint - projectedPoint);

            if (fallbackInteriorDirection == null)
            {
                XYZ apartmentInteriorPoint = GetApartmentInteriorReferencePoint(apartmentFi, doc);
                if (apartmentInteriorPoint != null)
                    fallbackInteriorDirection = Normalize2D(apartmentInteriorPoint - projectedPoint);
            }

            if (fallbackInteriorDirection == null)
                return false;

            interiorDirection = Dot2D(fallbackInteriorDirection, wallNormal) >= 0
                ? wallNormal
                : oppositeNormal;

            outwardDirection = new XYZ(-interiorDirection.X, -interiorDirection.Y, 0);
            exteriorSideSource = "внутренняя точка квартиры";
            return true;
        }

        private static bool IsPointInsideApartment2DRooms(FamilyInstance apartmentFi, Document doc, XYZ point)
        {
            if (point == null)
                return false;

            List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartmentFi);
            if (rooms == null || rooms.Count == 0)
                return false;

            double tolerance = IDHelper.ConvertMmToInternal(20);
            foreach (FamilyInstance roomFi in rooms)
            {
                if (IsPointInsideRoomGeometry2D(roomFi, point, tolerance))
                    return true;
            }

            return false;
        }

        private bool TryFindHostWallForDoorPlacement(Document doc, PreparedDoorPlacement preparedDoor, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, double maxDistanceToWallAxis, out Wall hostWall, out XYZ projectedPoint,
            out double distanceToWallAxis, out bool hostFromExistingWall)
        {
            hostWall = null;
            projectedPoint = null;
            distanceToWallAxis = double.MaxValue;
            hostFromExistingWall = false;

            if (preparedDoor == null || preparedDoor.InsertPoint == null)
                return false;

            if (TryFindBestHostWallForDoor(
                preparedDoor.InsertPoint,
                createdWallsForApartment,
                maxDistanceToWallAxis,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis))
            {
                hostFromExistingWall = false;
                return true;
            }

            List<Wall> existingWallCandidates = GetExistingWallsFromLineInfo(doc, existingWallsOnLevel);
            if (TryFindBestHostWallForDoor(
                preparedDoor.InsertPoint,
                existingWallCandidates,
                maxDistanceToWallAxis,
                true,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis))
            {
                hostFromExistingWall = true;
                return true;
            }

            return false;
        }

        private class EntranceDoorOrientationCandidate
        {
            public string Name { get; set; }
            public bool HandFlip { get; set; }
            public bool FacingFlip { get; set; }
            public double Score { get; set; }
            public XYZ HandDirection { get; set; }
            public XYZ FacingDirection { get; set; }
        }

        private static void OrientEntranceDoorBy2DSource(Document doc, FamilyInstance createdDoor, PreparedDoorPlacement preparedDoor,
            Wall hostWall, XYZ projectedDoorPoint, List<string> debugMessages, ApartmentProcessState state)
        {
            if (createdDoor == null || preparedDoor == null)
                return;

            XYZ toInterior = null;
            if (preparedDoor.InteriorReferencePoint != null && projectedDoorPoint != null)
                toInterior = Normalize2D(preparedDoor.InteriorReferencePoint - projectedDoorPoint);

            XYZ outward = null;
            if (preparedDoor.InteriorReferencePoint != null && projectedDoorPoint != null)
                outward = Normalize2D(projectedDoorPoint - preparedDoor.InteriorReferencePoint);

            XYZ wallNormal = GetWallAxisNormal2D(hostWall);

            XYZ initialHand = GetFamilyInstanceHandDirection2D(createdDoor, null);
            XYZ initialFacing = GetFamilyInstanceFacingDirection2D(createdDoor, null);

            bool canFlipHand = CanFlipHand(createdDoor);
            bool canFlipFacing = CanFlipFacing(createdDoor);

            List<EntranceDoorOrientationCandidate> candidates = new List<EntranceDoorOrientationCandidate>();
            bool handState = false;
            bool facingState = false;

            AddEntranceDoorOrientationCandidate(candidates, "без flip", handState, facingState, createdDoor, preparedDoor, toInterior);

            if (canFlipHand && TryFlipDoorHand(doc, createdDoor, debugMessages))
            {
                handState = true;
                AddEntranceDoorOrientationCandidate(candidates, "flipHand", handState, facingState, createdDoor, preparedDoor, toInterior);
            }

            if (canFlipFacing && TryFlipDoorFacing(doc, createdDoor, debugMessages))
            {
                facingState = true;
                AddEntranceDoorOrientationCandidate(
                    candidates,
                    handState ? "flipHand+flipFacing" : "flipFacing",
                    handState,
                    facingState,
                    createdDoor,
                    preparedDoor,
                    toInterior);
            }

            if (handState && TryFlipDoorHand(doc, createdDoor, debugMessages))
            {
                handState = false;
                AddEntranceDoorOrientationCandidate(
                    candidates,
                    facingState ? "flipFacing" : "без flip",
                    handState,
                    facingState,
                    createdDoor,
                    preparedDoor,
                    toInterior);
            }

            if (facingState && TryFlipDoorFacing(doc, createdDoor, debugMessages))
            {
                facingState = false;
                AddEntranceDoorOrientationCandidate(candidates, "без flip", handState, facingState, createdDoor, preparedDoor, toInterior);
            }

            EntranceDoorOrientationCandidate best = candidates
                .OrderByDescending(x => x.Score)
                .FirstOrDefault();

            if (best != null)
            {
                if (best.HandFlip)
                    TryFlipDoorHand(doc, createdDoor, debugMessages);

                if (best.FacingFlip)
                    TryFlipDoorFacing(doc, createdDoor, debugMessages);
            }

            OrientEntranceDoorOutside(createdDoor, preparedDoor.InteriorReferencePoint, projectedDoorPoint, debugMessages);

            XYZ finalHand = GetFamilyInstanceHandDirection2D(createdDoor, null);
            XYZ finalFacing = GetFamilyInstanceFacingDirection2D(createdDoor, null);

            AddApartmentDiagnostic(
                state,
                debugMessages,
                "Ориентация входной двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                ": наружу квартиры = " + FormatVector2D(outward) +
                ", нормаль оси стены = " + FormatVector2D(wallNormal) +
                ", 2D hand = " + FormatVector2D(preparedDoor.SourceHandDirection) +
                ", 2D facing = " + FormatVector2D(preparedDoor.SourceFacingDirection) +
                ", 3D было hand = " + FormatVector2D(initialHand) +
                ", facing = " + FormatVector2D(initialFacing) +
                ", canFlipHand = " + canFlipHand +
                ", canFlipFacing = " + canFlipFacing +
                ", выбран вариант = '" + (best != null ? best.Name : "<нет>") +
                "', 3D стало hand = " + FormatVector2D(finalHand) +
                ", facing = " + FormatVector2D(finalFacing) +
                ". Варианты: " + FormatEntranceDoorOrientationCandidates(candidates) + ".");
        }

        private static void AddEntranceDoorOrientationCandidate(List<EntranceDoorOrientationCandidate> candidates, string name, bool handFlip, bool facingFlip,
            FamilyInstance createdDoor, PreparedDoorPlacement preparedDoor, XYZ toInterior)
        {
            if (candidates == null || createdDoor == null || preparedDoor == null)
                return;

            XYZ hand = GetFamilyInstanceHandDirection2D(createdDoor, null);
            XYZ facing = GetFamilyInstanceFacingDirection2D(createdDoor, null);

            double score = 0;
            int sourceTerms = 0;

            if (preparedDoor.SourceHandDirection != null && hand != null)
            {
                score += Dot2D(preparedDoor.SourceHandDirection, hand);
                sourceTerms++;
            }

            if (preparedDoor.SourceFacingDirection != null && facing != null)
            {
                score += Dot2D(preparedDoor.SourceFacingDirection, facing);
                sourceTerms++;
            }

            if (preparedDoor.SourceHandDirection != null && preparedDoor.SourceFacingDirection != null && hand != null && facing != null)
            {
                double sourceCross = Cross2D(preparedDoor.SourceFacingDirection, preparedDoor.SourceHandDirection);
                double targetCross = Cross2D(facing, hand);

                if (Math.Abs(sourceCross) > 0.01 && Math.Abs(targetCross) > 0.01)
                    score += Math.Sign(sourceCross) == Math.Sign(targetCross) ? 1.5 : -1.5;
            }

            if (preparedDoor.IsEntranceDoor && toInterior != null && facing != null)
            {
                score += 6.0 * Dot2D(facing, new XYZ(-toInterior.X, -toInterior.Y, 0));
            }
            else if (toInterior != null && facing != null)
            {
                if (sourceTerms == 0)
                    score += Dot2D(facing, toInterior);
                else
                    score += 0.4 * Dot2D(facing, toInterior);
            }

            candidates.Add(new EntranceDoorOrientationCandidate
            {
                Name = name,
                HandFlip = handFlip,
                FacingFlip = facingFlip,
                Score = score,
                HandDirection = hand,
                FacingDirection = facing
            });
        }

        private static string FormatEntranceDoorOrientationCandidates(List<EntranceDoorOrientationCandidate> candidates)
        {
            if (candidates == null || candidates.Count == 0)
                return "<нет>";

            return string.Join("; ", candidates
                .Select(x =>
                    x.Name +
                    " score=" + Math.Round(x.Score, 3).ToString("0.###") +
                    " hand=" + FormatVector2D(x.HandDirection) +
                    " facing=" + FormatVector2D(x.FacingDirection))
                .Distinct()
                .ToArray());
        }

        private static bool TryFlipDoorHand(Document doc, FamilyInstance createdDoor, List<string> debugMessages)
        {
            if (createdDoor == null)
                return false;

            if (!CanFlipHand(createdDoor))
                return false;

            try
            {
                createdDoor.flipHand();
                if (doc != null)
                    doc.Regenerate();

                return true;
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось выполнить flipHand для входной двери ID = " + IDHelper.ElIdValue(createdDoor.Id) + ": " + ex.Message);

                return false;
            }
        }

        private static bool TryFlipDoorFacing(Document doc, FamilyInstance createdDoor, List<string> debugMessages)
        {
            if (createdDoor == null)
                return false;

            if (!CanFlipFacing(createdDoor))
                return false;

            try
            {
                createdDoor.flipFacing();
                if (doc != null)
                    doc.Regenerate();

                return true;
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось выполнить flipFacing для входной двери ID = " + IDHelper.ElIdValue(createdDoor.Id) + ": " + ex.Message);

                return false;
            }
        }

        private static void OrientEntranceDoorOutside(FamilyInstance createdDoor, XYZ interiorReferencePoint, XYZ projectedDoorPoint, List<string> debugMessages)
        {
            if (createdDoor == null || interiorReferencePoint == null || projectedDoorPoint == null)
                return;

            XYZ outward = Normalize2D(projectedDoorPoint - interiorReferencePoint);
            if (outward == null)
                return;

            XYZ facing;
            try
            {
                facing = Normalize2D(createdDoor.FacingOrientation);
            }
            catch
            {
                return;
            }

            if (facing == null || Dot2D(facing, outward) > 0)
                return;

            if (!CanFlipFacing(createdDoor))
            {
                if (debugMessages != null)
                    debugMessages.Add("Входная дверь ID = " + IDHelper.ElIdValue(createdDoor.Id) + " не поддерживает flipFacing для разворота наружу.");
                return;
            }

            try
            {
                createdDoor.flipFacing();
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось развернуть входную дверь ID = " + IDHelper.ElIdValue(createdDoor.Id) + " наружу: " + ex.Message);
            }
        }

        private PreparedApartmentDoors PrepareDoorsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, double placementPointZ,
            List<string> debugMessages, ApartmentProcessState state = null)
        {
            PreparedApartmentDoors result = new PreparedApartmentDoors();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyInstance> doorInstances = FindDoorSubComponentsRecursive(doc, apartmentFi);

            foreach (FamilyInstance doorFi in doorInstances)
            {
                if (doorFi == null)
                    continue;

                string typeName = doorFi.Symbol != null ? doorFi.Symbol.Name ?? "" : "";
                string commentValue = GetCommentsValue(doorFi);
                bool isEntranceDoor = HasEntranceDoorComment(doorFi);

                if (isEntranceDoor && state != null)
                    state.FoundEntranceDoorsCount++;

                int widthMm;
                if (!TryGetDoorWidthMmFrom2DMarker(doorFi, typeName, out widthMm) || widthMm <= 0)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не удалось определить ширину 2D-двери у экземпляра ID = " + IDHelper.ElIdValue(doorFi.Id));
                    continue;
                }

                Transform doorTransform = doorFi.GetTransform();
                if (doorTransform == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не удалось получить Transform для 2D-двери ID = " + IDHelper.ElIdValue(doorFi.Id));
                    continue;
                }

                XYZ insertPointInProject = WithZ(doorTransform.Origin, placementPointZ);
                XYZ sourceHandDirection = GetFamilyInstanceHandDirection2D(doorFi, doorTransform);
                XYZ sourceFacingDirection = GetFamilyInstanceFacingDirection2D(doorFi, doorTransform);

                TryAddPreparedDoorPlacement(
                    doc,
                    apartmentFi,
                    preset,
                    debugMessages,
                    state,
                    result,
                    doorFi,
                    doorFi.Id,
                    typeName,
                    commentValue,
                    isEntranceDoor,
                    widthMm,
                    insertPointInProject,
                    sourceHandDirection,
                    sourceFacingDirection,
                    false);
            }

            return result;
        }

        private bool TryAddPreparedDoorPlacement(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, List<string> debugMessages,
            ApartmentProcessState state, PreparedApartmentDoors result, FamilyInstance sourceDoorFi, ElementId sourceDoorId, string typeName, string commentValue, bool isEntranceDoor,
            int widthMm, XYZ insertPointInProject, XYZ sourceHandDirection, XYZ sourceFacingDirection, bool countEntranceDoor)
        {
            if (doc == null || apartmentFi == null || result == null || string.IsNullOrWhiteSpace(typeName) || widthMm <= 0 || insertPointInProject == null)
                return false;

            if (isEntranceDoor && countEntranceDoor && state != null)
                state.FoundEntranceDoorsCount++;

            string roomCategory = isEntranceDoor
                ? "Входная"
                : commentValue;

            if (string.IsNullOrWhiteSpace(roomCategory))
                roomCategory = "-";

            string presetKey = ApartmentDoorRequirementOption.BuildKey(roomCategory, typeName, widthMm, isEntranceDoor);

            string selectedDoorTypeName = null;
            if (preset != null && preset.DoorsByRoomCategory != null)
            {
                preset.DoorsByRoomCategory.TryGetValue(presetKey, out selectedDoorTypeName);

                if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                    string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                {
                    string legacyPresetKey = ApartmentDoorRequirementOption.BuildLegacyKey(roomCategory, typeName, widthMm, isEntranceDoor);
                    string legacySelectedDoorTypeName;
                    if (preset.DoorsByRoomCategory.TryGetValue(legacyPresetKey, out legacySelectedDoorTypeName) &&
                        !string.IsNullOrWhiteSpace(legacySelectedDoorTypeName) &&
                        !string.Equals(legacySelectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                    {
                        selectedDoorTypeName = legacySelectedDoorTypeName;
                    }
                }
            }

            List<string> entranceDiagnostics = new List<string>();
            if (isEntranceDoor)
            {
                entranceDiagnostics.Add(
                    "Входная 2D-дверь ID = " + FormatElementIdForDiagnostic(sourceDoorId) +
                    ": тип 2D = '" + typeName +
                    "', ширина = " + widthMm +
                    " мм, ключ пресета = '" + presetKey +
                    "', выбранный тип = '" + (string.IsNullOrWhiteSpace(selectedDoorTypeName) ? "<не выбран>" : selectedDoorTypeName) + "'.");
            }

            if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
            {
                string message = "Для двери [" + roomCategory + "] (" + widthMm + ") не выбран тип двери проекта.";
                if (isEntranceDoor)
                {
                    foreach (string diagnostic in entranceDiagnostics)
                        AddApartmentDiagnostic(state, debugMessages, diagnostic);

                    AddApartmentDiagnostic(state, debugMessages, message);
                }
                else if (debugMessages != null)
                    debugMessages.Add(message);

                return false;
            }

            FamilySymbol baseDoorSymbol = FindDoorSymbolByDisplayNameAndWidth(doc, selectedDoorTypeName, widthMm, isEntranceDoor);
            if (baseDoorSymbol == null)
            {
                string message = "Не найден тип двери проекта '" + selectedDoorTypeName + "' с шириной " + widthMm + " мм.";
                if (isEntranceDoor)
                {
                    foreach (string diagnostic in entranceDiagnostics)
                        AddApartmentDiagnostic(state, debugMessages, diagnostic);

                    AddApartmentDiagnostic(state, debugMessages, message);
                }
                else if (debugMessages != null)
                    debugMessages.Add(message);

                return false;
            }

            FamilySymbol resolvedDoorSymbol = sourceDoorFi == null
                ? baseDoorSymbol
                : ResolveDoorSymbolForPlacement(doc, typeName, baseDoorSymbol, debugMessages);

            if (resolvedDoorSymbol == null)
            {
                if (debugMessages != null)
                    debugMessages.Add(
                        "Не удалось определить итоговый тип 3D-двери для 2D-двери ID = " +
                        FormatElementIdForDiagnostic(sourceDoorId) + ".");
                return false;
            }

            FamilyInstance matchedRoom = FindBestMatchingRoomForDoor(
                apartmentFi,
                isEntranceDoor ? null : roomCategory,
                insertPointInProject,
                doc);

            XYZ expectedRoomPoint = matchedRoom != null
                ? GetRoomCenterPoint(matchedRoom)
                : GetApartmentInteriorReferencePoint(apartmentFi, doc);
            expectedRoomPoint = WithZ(expectedRoomPoint, insertPointInProject.Z);

            XYZ sourceRoomCalculationSideDirection = GetDoorRoomCalculationSideDirection(sourceDoorFi, insertPointInProject);

            if (isEntranceDoor)
            {
                entranceDiagnostics.Add(
                    "Входная дверь подготовлена: 2D ID = " + FormatElementIdForDiagnostic(sourceDoorId) +
                    ", итоговый тип = '" + BuildDoorTypeDisplayName(resolvedDoorSymbol) +
                    "', точка = " + FormatPointMm(insertPointInProject) +
                    ", внутренняя точка квартиры = " + FormatPointMm(expectedRoomPoint) +
                    ", 2D hand = " + FormatVector2D(sourceHandDirection) +
                    ", 2D facing = " + FormatVector2D(sourceFacingDirection) + ".");
            }

            result.Doors.Add(new PreparedDoorPlacement
            {
                ApartmentId = apartmentFi.Id,
                Door2DId = sourceDoorId,
                RoomCategory = roomCategory,
                DoorWidthMm = widthMm,
                SelectedDoorTypeName = BuildDoorTypeDisplayName(resolvedDoorSymbol),
                DoorSymbol = resolvedDoorSymbol,
                InsertPoint = insertPointInProject,
                RelatedRoom2D = matchedRoom,
                InteriorReferencePoint = expectedRoomPoint,
                SourceHandDirection = sourceHandDirection,
                SourceFacingDirection = sourceFacingDirection,
                SourceRoomCalculationSideDirection = sourceRoomCalculationSideDirection,
                IsEntranceDoor = isEntranceDoor,
                Diagnostics = entranceDiagnostics
            });

            return true;
        }

        private static DoorOpeningMarker GetDoorOpeningMarkerFromTypeName(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return DoorOpeningMarker.None;

            string name = typeName.Trim();

            if (name.IndexOf("(Л ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.Left;

            if (name.IndexOf("(Пр ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.RightAlt;

            if (name.IndexOf("(П ", StringComparison.OrdinalIgnoreCase) >= 0)
                return DoorOpeningMarker.Right;

            return DoorOpeningMarker.None;
        }

        private static string ReplaceDoorOpeningMarker(string typeName, DoorOpeningMarker newMarker)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return typeName;

            string replacement = "";

            switch (newMarker)
            {
                case DoorOpeningMarker.Left:
                    replacement = "(Л ";
                    break;
                case DoorOpeningMarker.Right:
                    replacement = "(П ";
                    break;
                case DoorOpeningMarker.RightAlt:
                    replacement = "(Пр ";
                    break;
                default:
                    return typeName;
            }

            string result = typeName;

            if (result.IndexOf("(Л ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(Л ", replacement);

            if (result.IndexOf("(Пр ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(Пр ", replacement);

            if (result.IndexOf("(П ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ReplaceOrdinalIgnoreCase(result, "(П ", replacement);

            return typeName;
        }

        private static string ReplaceOrdinalIgnoreCase(string source, string oldValue, string newValue)
        {
            int index = source.IndexOf(oldValue, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
                return source;

            return source.Substring(0, index) + newValue + source.Substring(index + oldValue.Length);
        }

        private static Parameter FindDoorLeftOpeningParameter(Element e)
        {
            if (e == null)
                return null;

            Parameter p = e.LookupParameter("КП_О_Левое открывание");
            if (p != null)
                return p;

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                p = typeElem.LookupParameter("КП_О_Левое открывание");
                if (p != null)
                    return p;
            }

            return null;
        }

        private static bool SetYesNoParameter(Parameter p, bool value)
        {
            if (p == null || p.IsReadOnly)
                return false;

            if (p.StorageType != StorageType.Integer)
                return false;

            p.Set(value ? 1 : 0);
            return true;
        }

        private static List<FamilySymbol> GetAllSymbolsOfFamily(Document doc, Family family)
        {
            List<FamilySymbol> result = new List<FamilySymbol>();
            if (doc == null || family == null)
                return result;

            ISet<ElementId> ids = family.GetFamilySymbolIds();
            if (ids == null || ids.Count == 0)
                return result;

            foreach (ElementId id in ids)
            {
                FamilySymbol symbol = doc.GetElement(id) as FamilySymbol;
                if (symbol != null)
                    result.Add(symbol);
            }

            return result;
        }

        private DoorTypeMirrorEnsureResult EnsureDoorMirrorTypeExists(Document doc, FamilySymbol sourceSymbol)
        {
            DoorTypeMirrorEnsureResult result = new DoorTypeMirrorEnsureResult();

            if (doc == null || sourceSymbol == null || sourceSymbol.Family == null)
                return result;

            string sourceTypeName = sourceSymbol.Name ?? "";
            DoorOpeningMarker marker = GetDoorOpeningMarkerFromTypeName(sourceTypeName);

            if (marker == DoorOpeningMarker.None)
                return result;

            string leftName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.Left);
            string rightName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.Right);
            string rightAltName = ReplaceDoorOpeningMarker(sourceTypeName, DoorOpeningMarker.RightAlt);

            Family family = sourceSymbol.Family;
            List<FamilySymbol> familySymbols = GetAllSymbolsOfFamily(doc, family);

            FamilySymbol existingLeft = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, leftName, StringComparison.OrdinalIgnoreCase));

            FamilySymbol existingRight = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, rightName, StringComparison.OrdinalIgnoreCase));

            FamilySymbol existingRightAlt = familySymbols.FirstOrDefault(x =>
                string.Equals(x.Name, rightAltName, StringComparison.OrdinalIgnoreCase));

            bool needCreateLeft = existingLeft == null;
            bool needCreateRight = existingRight == null && existingRightAlt == null;

            if (!needCreateLeft && !needCreateRight)
                return result;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Создание парных типов двери"))
            {
                t.Start();

                if (needCreateLeft)
                {
                    string newTypeName = leftName;
                    ElementType duplicated = sourceSymbol.Duplicate(newTypeName) as ElementType;
                    FamilySymbol newSymbol = duplicated as FamilySymbol;

                    if (newSymbol == null)
                        throw new Exception("Не удалось создать типоразмер '" + newTypeName + "'.");

                    Parameter p = FindDoorLeftOpeningParameter(newSymbol);
                    if (p == null)
                        throw new Exception("У нового типа '" + newTypeName + "' не найден параметр 'КП_О_Левое открывание'.");

                    if (!SetYesNoParameter(p, true))
                        throw new Exception("Не удалось включить параметр 'КП_О_Левое открывание' у типа '" + newTypeName + "'.");

                    result.HasMessage = true;
                    result.Message =
                        "Создан парный левый тип двери: '" + newTypeName + "'.";
                }

                if (needCreateRight)
                {
                    string newTypeName = rightName;
                    ElementType duplicated = sourceSymbol.Duplicate(newTypeName) as ElementType;
                    FamilySymbol newSymbol = duplicated as FamilySymbol;

                    if (newSymbol == null)
                        throw new Exception("Не удалось создать типоразмер '" + newTypeName + "'.");

                    Parameter p = FindDoorLeftOpeningParameter(newSymbol);
                    if (p == null)
                        throw new Exception("У нового типа '" + newTypeName + "' не найден параметр 'КП_О_Левое открывание'.");

                    if (!SetYesNoParameter(p, false))
                        throw new Exception("Не удалось выключить параметр 'КП_О_Левое открывание' у типа '" + newTypeName + "'.");

                    result.HasMessage = true;

                    if (string.IsNullOrWhiteSpace(result.Message))
                        result.Message = "Создан парный правый тип двери: '" + newTypeName + "'.";
                    else
                        result.Message += "\nСоздан парный правый тип двери: '" + newTypeName + "'.";
                }

                t.Commit();
            }

            return result;
        }

        private FamilySymbol ResolveDoorSymbolForPlacement(Document doc, string source2DTypeName, FamilySymbol baseDoorSymbol, List<string> debugMessages)
        {
            if (doc == null || string.IsNullOrWhiteSpace(source2DTypeName) || baseDoorSymbol == null)
                return baseDoorSymbol;

            string baseTypeName = baseDoorSymbol.Name ?? "";
            DoorOpeningMarker baseMarker = GetDoorOpeningMarkerFromTypeName(baseTypeName);

            if (baseMarker == DoorOpeningMarker.None)
                return baseDoorSymbol;

            DoorTypeMirrorEnsureResult ensureResult = EnsureDoorMirrorTypeExists(doc, baseDoorSymbol);
            if (ensureResult != null && ensureResult.HasMessage && debugMessages != null)
                debugMessages.Add(ensureResult.Message);

            DoorOpeningMarker desiredMarker;
            if (!TryGetDoorOpeningMarkerFrom2DTypeName(source2DTypeName, out desiredMarker))
                return baseDoorSymbol;

            if (desiredMarker == DoorOpeningMarker.None)
                return baseDoorSymbol;

            DoorOpeningMarker normalizedBaseMarker = NormalizeDoorOpeningMarkerForSelection(baseMarker);
            DoorOpeningMarker normalizedDesiredMarker = NormalizeDoorOpeningMarkerForSelection(desiredMarker);

            if (normalizedBaseMarker == normalizedDesiredMarker)
                return baseDoorSymbol;

            string desiredTypeName = ReplaceDoorOpeningMarker(baseTypeName, desiredMarker);

            FamilySymbol resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, desiredTypeName);
            if (resolved != null)
                return resolved;

            if (desiredMarker == DoorOpeningMarker.Right)
            {
                string rightAltName = ReplaceDoorOpeningMarker(baseTypeName, DoorOpeningMarker.RightAlt);
                resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, rightAltName);
                if (resolved != null)
                    return resolved;
            }

            if (desiredMarker == DoorOpeningMarker.RightAlt)
            {
                string rightName = ReplaceDoorOpeningMarker(baseTypeName, DoorOpeningMarker.Right);
                resolved = FindFamilySymbolByTypeName(doc, baseDoorSymbol.Family, rightName);
                if (resolved != null)
                    return resolved;
            }

            return baseDoorSymbol;
        }

        private static bool TryGetDoorOpeningMarkerFrom2DTypeName(string typeName, out DoorOpeningMarker marker)
        {
            marker = DoorOpeningMarker.None;

            if (string.IsNullOrWhiteSpace(typeName))
                return false;

            string normalized = typeName
                .Replace('ё', 'е')
                .Replace('Ё', 'Е')
                .Trim();

            for (int i = normalized.Length - 1; i >= 0; i--)
            {
                char c = normalized[i];
                if (char.IsWhiteSpace(c) || c == '.' || c == ',' || c == ';')
                    continue;

                if (c == 'Л' || c == 'л' || c == 'L' || c == 'l')
                {
                    marker = DoorOpeningMarker.Left;
                    return true;
                }

                if (c == 'П' || c == 'п' || c == 'P' || c == 'p')
                {
                    marker = DoorOpeningMarker.Right;
                    return true;
                }

                break;
            }

            return false;
        }

        private static FamilySymbol FindFamilySymbolByTypeName(Document doc, Family family, string typeName)
        {
            if (doc == null || family == null || string.IsNullOrWhiteSpace(typeName))
                return null;

            List<FamilySymbol> symbols = GetAllSymbolsOfFamily(doc, family);

            return symbols.FirstOrDefault(x =>
                string.Equals(x.Name, typeName, StringComparison.OrdinalIgnoreCase));
        }

        private static DoorOpeningMarker NormalizeDoorOpeningMarkerForSelection(DoorOpeningMarker marker)
        {
            if (marker == DoorOpeningMarker.RightAlt)
                return DoorOpeningMarker.Right;

            return marker;
        }

        private static List<FamilyInstance> FindDoorSubComponentsRecursive(Document doc, FamilyInstance rootInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (doc == null || rootInstance == null)
                return result;

            CollectDoorSubComponentsRecursive(doc, rootInstance, result);
            return result;
        }

        private static void CollectDoorSubComponentsRecursive(Document doc, FamilyInstance current, List<FamilyInstance> result)
        {
            if (doc == null || current == null)
                return;

            ICollection<ElementId> subIds = current.GetSubComponentIds();
            if (subIds == null || subIds.Count == 0)
                return;

            foreach (ElementId subId in subIds)
            {
                FamilyInstance subFi = doc.GetElement(subId) as FamilyInstance;
                if (subFi == null)
                    continue;

                string familyName = "";
                string typeName = "";
                string categoryName = "";

                if (subFi.Symbol != null)
                {
                    typeName = subFi.Symbol.Name ?? "";
                    if (subFi.Symbol.Family != null)
                        familyName = subFi.Symbol.Family.Name ?? "";
                }

                if (subFi.Category != null)
                    categoryName = subFi.Category.Name ?? "";

                if (Is2DDoorMarker(familyName, typeName, categoryName, HasEntranceDoorComment(subFi)))
                {
                    result.Add(subFi);
                    continue;
                }

                CollectDoorSubComponentsRecursive(doc, subFi, result);
            }
        }
    }
}