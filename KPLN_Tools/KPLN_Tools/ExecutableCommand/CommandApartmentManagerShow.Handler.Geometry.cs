using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private List<Line> BuildPreparedWallAxisLinesForSingleRoom(FamilyInstance roomFi, double apartmentWallThicknessInternal, List<ExistingWallLineInfo> existingWalls,
            double connectTol, double intersectionTol, List<string> debugMessages, ref int skippedWallsForApartment)
        {
            if (roomFi == null)
                return new List<Line>();

            CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi);

            List<Line> wallAxisLines = BuildClosedWallAxisLinesFromRooms(
                new List<CurveLoop> { roomLoop },
                apartmentWallThicknessInternal,
                debugMessages);

            if (wallAxisLines.Count == 0)
                return new List<Line>();

            List<Line> preparedAxisLines = SnapNewLinesToExistingWalls(wallAxisLines, existingWalls, connectTol);
            preparedAxisLines = MergeCollinearLines(preparedAxisLines);
            preparedAxisLines = RemoveSegmentsOverlappingExistingWalls(preparedAxisLines, existingWalls, apartmentWallThicknessInternal);
            preparedAxisLines = MergeCollinearLines(preparedAxisLines);

            if (preparedAxisLines.Count == 0)
                return new List<Line>();

            List<Line> finalAxisLines = new List<Line>();

            foreach (Line axis in preparedAxisLines)
            {
                if (axis == null || axis.Length < 1e-6)
                    continue;

                if (IntersectsExistingWalls(axis, apartmentWallThicknessInternal, existingWalls, intersectionTol))
                {
                    skippedWallsForApartment++;
                    continue;
                }

                finalAxisLines.Add(axis);
            }

            return MergeCollinearLines(finalAxisLines);
        }

        private static bool TryDeleteSource2DApartmentInstance(Document doc, ElementId apartmentId, List<string> debugMessages)
        {
            if (doc == null || apartmentId == null || apartmentId == ElementId.InvalidElementId)
                return false;

            try
            {
                Element element = doc.GetElement(apartmentId);
                if (element == null)
                    return false;

                doc.Delete(apartmentId);
                return true;
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось удалить 2D-квартиру ID = " + IDHelper.ElIdValue(apartmentId) + ": " + ex.Message);

                return false;
            }
        }

        private static bool IsFurnitureOrPlumbingCategory(Category category)
        {
            if (category == null)
                return false;

            BuiltInCategory bic;
            try
            {
                bic = (BuiltInCategory)IDHelper.ElIdInt(category.Id);
            }
            catch
            {
                return false;
            }

            return bic == BuiltInCategory.OST_Furniture ||
                   bic == BuiltInCategory.OST_PlumbingFixtures;
        }

        private static List<FamilyInstance> FindFurnitureAndPlumbingSubComponentsRecursive(Document doc, FamilyInstance rootInstance)
        {
            List<FamilyInstance> result = new List<FamilyInstance>();
            if (doc == null || rootInstance == null)
                return result;

            CollectFurnitureAndPlumbingSubComponentsRecursive(doc, rootInstance, result);
            return result;
        }

        private static void CollectFurnitureAndPlumbingSubComponentsRecursive(
            Document doc,
            FamilyInstance current,
            List<FamilyInstance> result)
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

                if (IsFurnitureOrPlumbingCategory(subFi.Category))
                {
                    result.Add(subFi);
                    continue;
                }

                CollectFurnitureAndPlumbingSubComponentsRecursive(doc, subFi, result);
            }
        }

        private static Level ResolvePlacementLevelForNestedInstance(Document doc, FamilyInstance nestedFi, FamilyInstance apartmentFi)
        {
            if (doc == null)
                return null;

            ElementId nestedLevelId = GetInstanceLevelId(nestedFi);
            if (nestedLevelId != ElementId.InvalidElementId)
            {
                Level nestedLevel = doc.GetElement(nestedLevelId) as Level;
                if (nestedLevel != null)
                    return nestedLevel;
            }

            ElementId apartmentLevelId = GetInstanceLevelId(apartmentFi);
            if (apartmentLevelId != ElementId.InvalidElementId)
            {
                Level apartmentLevel = doc.GetElement(apartmentLevelId) as Level;
                if (apartmentLevel != null)
                    return apartmentLevel;
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .FirstOrDefault();
        }

        private static double GetRotationAngleOnXY(FamilyInstance fi)
        {
            if (fi == null)
                return 0.0;

            Transform tr = fi.GetTransform();
            if (tr == null)
                return 0.0;

            XYZ basisX = tr.BasisX;
            if (basisX == null)
                return 0.0;

            return Math.Atan2(basisX.Y, basisX.X);
        }

        private void CopyFurnitureAndPlumbingFromApartmentUnderlay(Document doc, FamilyInstance apartmentFi, List<string> debugMessages, List<ElementId> createdElementIds = null)
        {
            if (doc == null || apartmentFi == null)
                return;

            List<FamilyInstance> nestedItems = FindFurnitureAndPlumbingSubComponentsRecursive(doc, apartmentFi);
            if (nestedItems == null || nestedItems.Count == 0)
                return;

            foreach (FamilyInstance nestedFi in nestedItems)
            {
                if (nestedFi == null)
                    continue;

                try
                {
                    FamilySymbol symbol = nestedFi.Symbol;
                    if (symbol == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + " не найден тип.");
                        continue;
                    }

                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                        doc.Regenerate();
                    }

                    Transform tr = nestedFi.GetTransform();
                    if (tr == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("У вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + " не найден Transform.");
                        continue;
                    }

                    XYZ insertPoint = tr.Origin;
                    if (insertPoint == null)
                        continue;

                    Level level = ResolvePlacementLevelForNestedInstance(doc, nestedFi, apartmentFi);
                    if (level == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не найден уровень для вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id));
                        continue;
                    }

                    FamilyPlacementType placementType = symbol.Family.FamilyPlacementType;
                    FamilyInstance created = null;

                    switch (placementType)
                    {
                        case FamilyPlacementType.ViewBased:
                            break;

                        case FamilyPlacementType.OneLevelBased:
                        case FamilyPlacementType.OneLevelBasedHosted:
                        case FamilyPlacementType.WorkPlaneBased:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;

                        default:
                            created = doc.Create.NewFamilyInstance(
                                insertPoint,
                                symbol,
                                level,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            break;
                    }

                    if (created == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не удалось создать экземпляр для вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id));
                        continue;
                    }

                    if (createdElementIds != null)
                        createdElementIds.Add(created.Id);

                    double angle = GetRotationAngleOnXY(nestedFi);
                    if (Math.Abs(angle) > 1e-9)
                    {
                        Line axis = Line.CreateBound(insertPoint, insertPoint + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, created.Id, axis, angle);
                    }

                    ApplyFamilyInstanceFlipState(nestedFi, created, debugMessages);
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Ошибка копирования вложенного элемента ID = " + IDHelper.ElIdValue(nestedFi.Id) + ": " + ex.Message);
                }
            }
        }

        private static void ApplyFamilyInstanceFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            if (source == null || target == null)
                return;

            ApplyHandFlipState(source, target, debugMessages);
            ApplyFacingFlipState(source, target, debugMessages);
        }

        private static void ApplyHandFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            bool sourceValue;
            bool targetValue;

            if (!TryGetHandFlipped(source, out sourceValue))
                return;

            if (!TryGetHandFlipped(target, out targetValue))
                return;

            if (sourceValue == targetValue)
                return;

            if (!CanFlipHand(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно повторить flipHand у вложенного элемента ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipHand.");
                return;
            }

            try
            {
                target.flipHand();
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось применить flipHand к вложенному элементу ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static void ApplyFacingFlipState(FamilyInstance source, FamilyInstance target, List<string> debugMessages)
        {
            bool sourceValue;
            bool targetValue;

            if (!TryGetFacingFlipped(source, out sourceValue))
                return;

            if (!TryGetFacingFlipped(target, out targetValue))
                return;

            if (sourceValue == targetValue)
                return;

            if (!CanFlipFacing(target))
            {
                if (debugMessages != null)
                    debugMessages.Add("Невозможно повторить flipFacing у вложенного элемента ID = " + IDHelper.ElIdValue(source.Id) + ": созданный тип не поддерживает flipFacing.");
                return;
            }

            try
            {
                target.flipFacing();
            }
            catch (Exception ex)
            {
                if (debugMessages != null)
                    debugMessages.Add("Не удалось применить flipFacing к вложенному элементу ID = " + IDHelper.ElIdValue(source.Id) + ": " + ex.Message);
            }
        }

        private static bool TryGetHandFlipped(FamilyInstance fi, out bool value)
        {
            value = false;

            if (fi == null)
                return false;

            try
            {
                value = fi.HandFlipped;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetFacingFlipped(FamilyInstance fi, out bool value)
        {
            value = false;

            if (fi == null)
                return false;

            try
            {
                value = fi.FacingFlipped;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool CanFlipHand(FamilyInstance fi)
        {
            if (fi == null)
                return false;

            try
            {
                return fi.CanFlipHand;
            }
            catch
            {
                return false;
            }
        }

        private static bool CanFlipFacing(FamilyInstance fi)
        {
            if (fi == null)
                return false;

            try
            {
                return fi.CanFlipFacing;
            }
            catch
            {
                return false;
            }
        }

        private static XYZ GetFamilyInstanceHandDirection2D(FamilyInstance fi, Transform fallbackTransform)
        {
            if (fi != null)
            {
                try
                {
                    XYZ hand = fi.HandOrientation;
                    if (hand != null)
                    {
                        XYZ normalized = Normalize2D(hand);
                        if (normalized != null)
                            return normalized;
                    }
                }
                catch
                {
                }
            }

            if (fallbackTransform == null && fi != null)
            {
                try
                {
                    fallbackTransform = fi.GetTransform();
                }
                catch
                {
                }
            }

            if (fallbackTransform != null)
            {
                XYZ basisX = fallbackTransform.BasisX;
                if (basisX != null)
                    return Normalize2D(basisX);
            }

            return null;
        }

        private static XYZ GetFamilyInstanceFacingDirection2D(FamilyInstance fi, Transform fallbackTransform)
        {
            if (fi != null)
            {
                try
                {
                    XYZ facing = fi.FacingOrientation;
                    if (facing != null)
                    {
                        XYZ normalized = Normalize2D(facing);
                        if (normalized != null)
                            return normalized;
                    }
                }
                catch
                {
                }
            }

            if (fallbackTransform == null && fi != null)
            {
                try
                {
                    fallbackTransform = fi.GetTransform();
                }
                catch
                {
                }
            }

            if (fallbackTransform != null)
            {
                XYZ basisY = fallbackTransform.BasisY;
                if (basisY != null)
                    return Normalize2D(basisY);
            }

            return null;
        }

        private static void DeleteIsolatedCreatedWalls(Document doc, List<Wall> createdWallsForApartment)
        {
            if (doc == null || createdWallsForApartment == null || createdWallsForApartment.Count == 0)
                return;

            double tol = IDHelper.ConvertMmToInternal(10);
            List<Wall> validWalls = createdWallsForApartment
                .Where(x => x != null && x.IsValidObject)
                .ToList();

            if (validWalls.Count == 0)
                return;

            List<ElementId> wallsToDelete = new List<ElementId>();

            foreach (Wall currentWall in validWalls)
            {
                if (currentWall == null || !currentWall.IsValidObject)
                    continue;

                LocationCurve currentLc = currentWall.Location as LocationCurve;
                if (currentLc == null)
                    continue;

                Line currentLine = currentLc.Curve as Line;
                if (currentLine == null || currentLine.Length < 1e-9)
                    continue;

                XYZ a0 = currentLine.GetEndPoint(0);
                XYZ a1 = currentLine.GetEndPoint(1);

                bool hasConnection = false;

                foreach (Wall otherWall in validWalls)
                {
                    if (otherWall == null || !otherWall.IsValidObject || otherWall.Id == currentWall.Id)
                        continue;

                    LocationCurve otherLc = otherWall.Location as LocationCurve;
                    if (otherLc == null)
                        continue;

                    Line otherLine = otherLc.Curve as Line;
                    if (otherLine == null || otherLine.Length < 1e-9)
                        continue;

                    XYZ b0 = otherLine.GetEndPoint(0);
                    XYZ b1 = otherLine.GetEndPoint(1);

                    bool touchesByEndpoints =
                        Distance2D(a0, b0) <= tol ||
                        Distance2D(a0, b1) <= tol ||
                        Distance2D(a1, b0) <= tol ||
                        Distance2D(a1, b1) <= tol;

                    XYZ intersection;
                    bool intersects = TryIntersectSegments2D(a0, a1, b0, b1, out intersection, tol);

                    if (touchesByEndpoints || intersects)
                    {
                        hasConnection = true;
                        break;
                    }
                }

                if (!hasConnection)
                    wallsToDelete.Add(currentWall.Id);
            }

            if (wallsToDelete.Count > 0)
            {
                doc.Delete(wallsToDelete);

                createdWallsForApartment.RemoveAll(x =>
                    x == null ||
                    !x.IsValidObject ||
                    wallsToDelete.Any(id => id == x.Id));
            }
        }

        private static XYZ GetRoomCenterPoint(FamilyInstance roomFi)
        {
            if (roomFi == null)
                return null;

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                return null;

            return tr.Origin;
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

        private static Line GetWallAxisLine(Wall wall)
        {
            if (wall == null)
                return null;

            LocationCurve lc = wall.Location as LocationCurve;
            if (lc == null)
                return null;

            return lc.Curve as Line;
        }

        private static XYZ GetClosestPointOnRoomRectangle(FamilyInstance roomFi, XYZ worldPoint)
        {
            if (roomFi == null || worldPoint == null)
                return null;

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                return null;

            Transform inv = tr.Inverse;
            if (inv == null)
                return null;

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            XYZ localPoint = inv.OfPoint(worldPoint);

            double clampedX = Math.Max(-halfW, Math.Min(halfW, localPoint.X));
            double clampedY = Math.Max(-halfD, Math.Min(halfD, localPoint.Y));

            XYZ localClosest = new XYZ(clampedX, clampedY, 0);
            return tr.OfPoint(localClosest);
        }

        private void CorrectWallDirectionsForApartmentBy2DDoors(
            Document doc,
            PreparedApartmentDoors apartmentDoors,
            List<Wall> createdWallsForApartment)
        {
            if (doc == null || apartmentDoors == null || apartmentDoors.Doors == null || createdWallsForApartment == null)
                return;

            double maxDistanceToWallAxis = IDHelper.ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                if (preparedDoor == null)
                    continue;

                preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = false;

                if (preparedDoor.IsEntranceDoor)
                    continue;

                if (preparedDoor.InsertPoint == null || preparedDoor.RelatedRoom2D == null)
                    continue;

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;

                bool foundHost = TryFindBestHostWallForDoor(
                    preparedDoor.InsertPoint,
                    createdWallsForApartment,
                    maxDistanceToWallAxis,
                    out hostWall,
                    out projectedPoint,
                    out distanceToWallAxis);

                if (!foundHost || hostWall == null || projectedPoint == null)
                    continue;

                Line wallAxis = GetWallAxisLine(hostWall);
                if (wallAxis == null)
                    continue;

                XYZ wallDir = Normalize2D(wallAxis.GetEndPoint(1) - wallAxis.GetEndPoint(0));
                if (wallDir == null)
                    continue;

                XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0);

                XYZ roomPoint = GetClosestPointOnRoomRectangle(preparedDoor.RelatedRoom2D, preparedDoor.InsertPoint);
                if (roomPoint == null)
                    continue;

                XYZ toRoom = roomPoint - projectedPoint;
                double sign = Dot2D(toRoom, wallNormal);


                if (sign > 0)
                {
                    preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = true;

                    LocationCurve lc = hostWall.Location as LocationCurve;
                    if (lc == null)
                        continue;

                    Line reversedAxis = Line.CreateBound(
                        wallAxis.GetEndPoint(1),
                        wallAxis.GetEndPoint(0));

                    lc.Curve = reversedAxis;
                }
            }
        }

        private int PlaceDoorsForApartment(Document doc, PreparedApartmentDoors apartmentDoors, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, Level baseLevel, List<string> debugMessages, ApartmentProcessState state,
            List<ElementId> createdDoorIds = null)
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
                        createdDoorIds))
                    {
                        installedCount++;
                    }

                    continue;
                }

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;
                bool hostFromExistingWall;

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

                bool useOppositeDoorType =
                    preparedDoor.RequiresOppositeDoorTypeAfterWallFlip ||
                    (hostFromExistingWall &&
                     ShouldUseOppositeDoorTypeForExistingHost(hostWall, preparedDoor, projectedPoint));

                if (useOppositeDoorType)
                {
                    symbolToPlace = GetOppositeDoorSymbol(doc, symbolToPlace);
                }

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

                try
                {
                    createdDoor = CreateDoorFamilyInstance(
                        doc,
                        projectedPoint,
                        symbolToPlace,
                        hostWall,
                        baseLevel,
                        preparedDoor,
                        debugMessages);
                }
                catch (Exception ex)
                {
                    AddPreparedDoorDiagnostics(state, debugMessages, preparedDoor);

                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Ошибка вставки двери ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                        " в стену ID = " + IDHelper.ElIdValue(hostWall.Id) +
                        ", тип '" + BuildDoorTypeDisplayName(symbolToPlace) + "': " + ex.Message);
                    continue;
                }

                if (createdDoor != null)
                {
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

        private bool PlaceEntranceDoorForApartment(Document doc, PreparedDoorPlacement preparedDoor, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, Level baseLevel, List<string> debugMessages, ApartmentProcessState state,
            List<ElementId> createdDoorIds = null)
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
            preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = false;
            List<string> entranceDiagnostics = new List<string>();

            bool hostAxisReversed = EnsureEntranceHostWallNormalOutside(
                doc,
                hostWall,
                outwardDirection,
                preparedDoor,
                !hostFromExistingWall,
                entranceDiagnostics,
                null);

            FamilySymbol sourceEntranceDoorSymbol = preparedDoor.DoorSymbol;
            FamilySymbol entranceDoorSymbol = ResolveEntranceDoorSymbolForExterior(doc, preparedDoor, sourceEntranceDoorSymbol, outwardDirection);
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

            if (sourceEntranceDoorSymbol != null &&
                entranceDoorSymbol != null &&
                IDHelper.ElIdValue(sourceEntranceDoorSymbol.Id) != IDHelper.ElIdValue(entranceDoorSymbol.Id))
            {
                AddApartmentDiagnostic(
                    null,
                    entranceDiagnostics,
                    "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor.Door2DId) +
                    ": 2D facing смотрит внутрь квартиры, выбран парный тип '" +
                    BuildDoorTypeDisplayName(entranceDoorSymbol) + "' вместо '" +
                    BuildDoorTypeDisplayName(sourceEntranceDoorSymbol) + "'.");
            }

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
                ", ось хоста развернута = " + hostAxisReversed +
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
                        sourceEntranceDoorSymbol,
                        entranceDoorSymbol,
                        entranceReferenceDirection,
                        useEntranceReferenceDirection,
                        createdWallsForApartment,
                        createdDoorIds,
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

            if (createdDoorIds != null)
                createdDoorIds.Add(createdDoor.Id);

            return true;
        }

        private int PlaceWindowsForApartment(Document doc, PreparedApartmentWindows apartmentWindows, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, Level baseLevel, List<string> debugMessages, ApartmentProcessState state,
            List<ElementId> createdWindowIds = null)
        {
            if (doc == null || apartmentWindows == null || apartmentWindows.Windows == null || apartmentWindows.Windows.Count == 0 || baseLevel == null)
                return 0;

            int installedCount = 0;
            double maxDistanceToWallAxis = IDHelper.ConvertMmToInternal(2000);

            foreach (PreparedWindowPlacement preparedWindow in apartmentWindows.Windows)
            {
                if (preparedWindow == null || preparedWindow.InsertPoint == null || preparedWindow.WindowSymbol == null)
                    continue;

                Wall hostWall;
                XYZ projectedPoint;
                double distanceToWallAxis;
                bool hostFromExistingWall;

                bool foundHost = TryFindHostWallForWindowPlacement(
                    doc,
                    preparedWindow,
                    createdWallsForApartment,
                    existingWallsOnLevel,
                    maxDistanceToWallAxis,
                    out hostWall,
                    out projectedPoint,
                    out distanceToWallAxis,
                    out hostFromExistingWall);

                if (!foundHost || hostWall == null || projectedPoint == null)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не найдена стена-хост для окна квартиры ID = " +
                        IDHelper.ElIdValue(preparedWindow.ApartmentId) +
                        ". Точка = " + FormatPointMm(preparedWindow.InsertPoint) +
                        ", линия = " + FormatPointMm(preparedWindow.SourceLine != null ? preparedWindow.SourceLine.GetEndPoint(0) : null) +
                        " -> " + FormatPointMm(preparedWindow.SourceLine != null ? preparedWindow.SourceLine.GetEndPoint(1) : null) +
                        ". Ближайшая созданная: " + BuildNearestWallDiagnostic(preparedWindow.InsertPoint, createdWallsForApartment, false, maxDistanceToWallAxis) +
                        ". Ближайшая существующая: " + BuildNearestWallDiagnostic(preparedWindow.InsertPoint, GetExistingWallsFromLineInfo(doc, existingWallsOnLevel), true, maxDistanceToWallAxis));
                    continue;
                }

                FamilySymbol symbolToPlace = preparedWindow.WindowSymbol;

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
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не удалось активировать тип окна '" + BuildWindowTypeDisplayName(symbolToPlace) +
                        "' для квартиры ID = " + IDHelper.ElIdValue(preparedWindow.ApartmentId) + ": " + ex.Message);
                    continue;
                }

                FamilyInstance createdWindow = null;

                try
                {
                    createdWindow = CreateWindowFamilyInstance(
                        doc,
                        projectedPoint,
                        symbolToPlace,
                        hostWall,
                        baseLevel,
                        preparedWindow.ReferenceDirection,
                        debugMessages);
                }
                catch (Exception ex)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Ошибка вставки окна квартиры ID = " + IDHelper.ElIdValue(preparedWindow.ApartmentId) +
                        " в стену ID = " + IDHelper.ElIdValue(hostWall.Id) +
                        ", тип '" + BuildWindowTypeDisplayName(symbolToPlace) + "': " + ex.Message);
                    continue;
                }

                if (createdWindow == null)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Revit не создал окно квартиры ID = " + IDHelper.ElIdValue(preparedWindow.ApartmentId) +
                        " в стене ID = " + IDHelper.ElIdValue(hostWall.Id) +
                        ", тип '" + BuildWindowTypeDisplayName(symbolToPlace) + "'.");
                    continue;
                }

                if (!TrySetWindowSillHeight(createdWindow, preparedWindow.SillHeightInternal))
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Окно ID = " + IDHelper.ElIdValue(createdWindow.Id) +
                        " создано, но не удалось задать параметр 'Высота нижнего бруса'.");
                }

                installedCount++;

                if (createdWindowIds != null)
                    createdWindowIds.Add(createdWindow.Id);
            }

            return installedCount;
        }

        private static FamilyInstance CreateWindowFamilyInstance(Document doc, XYZ projectedPoint, FamilySymbol symbolToPlace, Wall hostWall,
            Level baseLevel, XYZ referenceDirection, List<string> debugMessages)
        {
            if (doc == null || projectedPoint == null || symbolToPlace == null || hostWall == null)
                return null;

            XYZ refDir = referenceDirection;
            if (refDir == null)
                refDir = GetWallAxisDirection2D(hostWall);

            if (refDir != null)
            {
                try
                {
                    return doc.Create.NewFamilyInstance(
                        projectedPoint,
                        symbolToPlace,
                        refDir,
                        hostWall,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Вставка окна с referenceDirection не сработала, используется стандартная вставка: " + ex.Message);
                }
            }

            return doc.Create.NewFamilyInstance(
                projectedPoint,
                symbolToPlace,
                hostWall,
                baseLevel,
                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }

        private static bool TrySetWindowSillHeight(FamilyInstance window, double sillHeightInternal)
        {
            if (window == null || sillHeightInternal <= 0)
                return false;

            Parameter p = window.LookupParameter("Высота нижнего бруса");
            if (p == null)
                p = window.LookupParameter("Sill Height");

            if (p == null)
            {
                try
                {
                    p = window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
                }
                catch
                {
                    p = null;
                }
            }

            if (p == null || p.IsReadOnly || p.StorageType != StorageType.Double)
                return false;

            try
            {
                p.Set(sillHeightInternal);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryFindHostWallForWindowPlacement(Document doc, PreparedWindowPlacement preparedWindow, List<Wall> createdWallsForApartment,
            List<ExistingWallLineInfo> existingWallsOnLevel, double maxDistanceToWallAxis, out Wall hostWall, out XYZ projectedPoint,
            out double distanceToWallAxis, out bool hostFromExistingWall)
        {
            hostWall = null;
            projectedPoint = null;
            distanceToWallAxis = double.MaxValue;
            hostFromExistingWall = false;

            if (preparedWindow == null || preparedWindow.InsertPoint == null)
                return false;

            if (TryFindBestHostWallForWindow(
                preparedWindow,
                createdWallsForApartment,
                maxDistanceToWallAxis,
                false,
                out hostWall,
                out projectedPoint,
                out distanceToWallAxis))
            {
                hostFromExistingWall = false;
                return true;
            }

            List<Wall> existingWallCandidates = GetExistingWallsFromLineInfo(doc, existingWallsOnLevel);
            if (TryFindBestHostWallForWindow(
                preparedWindow,
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

        private bool TryFindBestHostWallForWindow(PreparedWindowPlacement preparedWindow, List<Wall> candidateWalls, double maxDistanceToWallAxis,
            bool includeWallHalfWidth, out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            bestWall = null;
            bestProjectedPoint = null;
            bestDistance = double.MaxValue;
            double bestScore = double.MaxValue;

            if (preparedWindow == null || preparedWindow.InsertPoint == null || candidateWalls == null || candidateWalls.Count == 0)
                return false;

            XYZ windowDir = preparedWindow.ReferenceDirection;

            foreach (Wall wall in candidateWalls)
            {
                if (wall == null)
                    continue;

                Line wallLine = GetWallAxisLine(wall);
                if (wallLine == null)
                    continue;

                XYZ wallDir = Normalize2D(wallLine.GetEndPoint(1) - wallLine.GetEndPoint(0));
                if (wallDir == null)
                    continue;

                double parallelPenalty = 0;
                if (windowDir != null)
                {
                    double parallel = Math.Abs(Dot2D(windowDir, wallDir));
                    if (parallel < 0.80)
                        continue;

                    parallelPenalty = (1.0 - parallel) * IDHelper.ConvertMmToInternal(500);
                }

                XYZ projectedPoint;
                double distance;
                if (!TryProjectPointToSegment2D(preparedWindow.InsertPoint, wallLine, out projectedPoint, out distance))
                    continue;

                double wallHalfWidth = includeWallHalfWidth ? GetWallHalfWidth(wall) : 0;
                double allowedDistance = maxDistanceToWallAxis + wallHalfWidth;

                if (distance > allowedDistance)
                    continue;

                double score = Math.Max(0, distance - wallHalfWidth) + parallelPenalty;
                if (score < bestScore)
                {
                    bestScore = score;
                    bestDistance = distance;
                    bestWall = wall;
                    bestProjectedPoint = projectedPoint;
                }
            }

            return bestWall != null;
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

        private static string FormatPointMm(XYZ point)
        {
            if (point == null)
                return "<нет>";

            return "(" +
                   Math.Round(IDHelper.ConvertInternalToMm(point.X)).ToString("0") + "; " +
                   Math.Round(IDHelper.ConvertInternalToMm(point.Y)).ToString("0") + "; " +
                   Math.Round(IDHelper.ConvertInternalToMm(point.Z)).ToString("0") + ") мм";
        }

        private static string FormatLengthMm(double valueInternal)
        {
            return Math.Round(IDHelper.ConvertInternalToMm(valueInternal)).ToString("0") + " мм";
        }

        private static string FormatDouble(double value)
        {
            if (double.IsNaN(value))
                return "<nan>";

            if (double.IsNegativeInfinity(value) || double.IsPositiveInfinity(value))
                return "<нет>";

            return Math.Round(value, 3).ToString("0.###");
        }

        private static void AddApartmentDiagnostics(ApartmentProcessState state, List<string> debugMessages, IEnumerable<string> messages)
        {
            if (messages == null)
                return;

            foreach (string message in messages)
                AddApartmentDiagnostic(state, debugMessages, message);
        }

        private static string FormatVector2D(XYZ vector)
        {
            if (vector == null)
                return "<нет>";

            return "(" +
                   Math.Round(vector.X, 3).ToString("0.###") + "; " +
                   Math.Round(vector.Y, 3).ToString("0.###") + ")";
        }

        private static bool ShouldUseOppositeDoorTypeForExistingHost(Wall hostWall, PreparedDoorPlacement preparedDoor, XYZ projectedPoint)
        {
            if (hostWall == null || preparedDoor == null || projectedPoint == null)
                return false;

            XYZ referencePoint = null;

            if (preparedDoor.RelatedRoom2D != null)
            {
                referencePoint = GetClosestPointOnRoomRectangle(preparedDoor.RelatedRoom2D, preparedDoor.InsertPoint) ??
                                 GetRoomCenterPoint(preparedDoor.RelatedRoom2D);
            }

            if (referencePoint == null)
                referencePoint = preparedDoor.InteriorReferencePoint;

            if (referencePoint == null)
                return false;

            Line wallAxis = GetWallAxisLine(hostWall);
            if (wallAxis == null)
                return false;

            XYZ wallDir = Normalize2D(wallAxis.GetEndPoint(1) - wallAxis.GetEndPoint(0));
            if (wallDir == null)
                return false;

            XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0);
            XYZ toReference = referencePoint - projectedPoint;

            return Dot2D(toReference, wallNormal) > 0;
        }

        private static XYZ GetWallAxisNormal2D(Wall wall)
        {
            XYZ wallDir = GetWallAxisDirection2D(wall);
            if (wallDir == null)
                return null;

            return new XYZ(-wallDir.Y, wallDir.X, 0);
        }

        private static XYZ GetWallAxisDirection2D(Wall wall)
        {
            Line wallAxis = GetWallAxisLine(wall);
            if (wallAxis == null)
                return null;

            return Normalize2D(wallAxis.GetEndPoint(1) - wallAxis.GetEndPoint(0));
        }

        private static FamilyInstance CreateDoorFamilyInstance(Document doc, XYZ projectedPoint, FamilySymbol symbolToPlace, Wall hostWall,
            Level baseLevel, PreparedDoorPlacement preparedDoor, List<string> debugMessages)
        {
            if (doc == null || projectedPoint == null || symbolToPlace == null || hostWall == null)
                return null;

            return doc.Create.NewFamilyInstance(
                projectedPoint,
                symbolToPlace,
                hostWall,
                baseLevel,
                Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }

        private FamilySymbol ResolveEntranceDoorSymbolForExterior(Document doc, PreparedDoorPlacement preparedDoor, FamilySymbol baseSymbol,
            XYZ outwardDirection)
        {
            if (doc == null || preparedDoor == null || baseSymbol == null)
                return baseSymbol;

            XYZ sourceFacing = Normalize2D(preparedDoor.SourceFacingDirection);
            XYZ outward = Normalize2D(outwardDirection);

            if (sourceFacing == null || outward == null)
                return baseSymbol;

            double facingToOutside = Dot2D(sourceFacing, outward);
            if (facingToOutside >= -0.25)
                return baseSymbol;

            FamilySymbol oppositeSymbol = GetOppositeDoorSymbol(doc, baseSymbol);
            if (oppositeSymbol == null || IDHelper.ElIdValue(oppositeSymbol.Id) == IDHelper.ElIdValue(baseSymbol.Id))
                return baseSymbol;

            return oppositeSymbol;
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

            if (!allowReverseHostAxis)
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor != null ? preparedDoor.Door2DId : ElementId.InvalidElementId) +
                    ": существующая стена-хост не разворачивается. Нормаль стены = " +
                    FormatVector2D(wallNormal) + ", наружу квартиры = " + FormatVector2D(outward) + ".");

                return false;
            }

            try
            {
                LocationCurve lc = hostWall.Location as LocationCurve;
                Line wallAxis = GetWallAxisLine(hostWall);
                if (lc == null || wallAxis == null)
                    return false;

                lc.Curve = Line.CreateBound(
                    wallAxis.GetEndPoint(1),
                    wallAxis.GetEndPoint(0));

                if (doc != null)
                    doc.Regenerate();

                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Входная дверь ID = " + FormatElementIdForDiagnostic(preparedDoor != null ? preparedDoor.Door2DId : ElementId.InvalidElementId) +
                    ": ось стены-хоста развернута перед вставкой, чтобы нормаль стены смотрела наружу квартиры. Было = " +
                    FormatVector2D(wallNormal) + ", наружу = " + FormatVector2D(outward) +
                    ", стало = " + FormatVector2D(GetWallAxisNormal2D(hostWall)) + ".");

                return true;
            }
            catch (Exception ex)
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Не удалось развернуть ось стены-хоста для входной двери ID = " +
                    FormatElementIdForDiagnostic(preparedDoor != null ? preparedDoor.Door2DId : ElementId.InvalidElementId) +
                    ": " + ex.Message);

                return false;
            }
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
            if (useReferenceDirection)
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
            List<string> diagnostics, out string failureDiagnostic)
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
            List<EntranceDoorCreationCandidate> candidates, List<string> attempts)
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

            Line auxiliaryAxis = BuildEntranceAuxiliaryHostWallAxis(projectedPoint, existingHostWall, outwardDirection, hostWallLength);
            if (auxiliaryAxis == null)
            {
                if (attempts != null)
                    attempts.Add("fallback короткая стена-хост: не удалось построить ось");
                return null;
            }

            double wallHeight = GetWallHeightInternal(existingHostWall);
            if (wallHeight <= IDHelper.ConvertMmToInternal(500))
                wallHeight = IDHelper.ConvertMmToInternal(3000);

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
                ApplyWallPresetParameters(auxiliaryWall, baseLevel, null, 0, wallHeight);
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
            AddUniqueFamilySymbol(symbols, GetOppositeDoorSymbol(doc, preferredSymbol));
            AddUniqueFamilySymbol(symbols, GetOppositeDoorSymbol(doc, sourceSymbol));

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

            double z0 = baseLevel.Elevation;
            double z1 = z0 + doorHeight + IDHelper.ConvertMmToInternal(150);
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
                    diagnostics.Add("Не удалось удалить неподходящий вариант входной двери ID = " + IDHelper.ElIdValue(elementId) + ": " + ex.Message);

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

            try
            {
                return doc.Create.NewFamilyInstance(
                    projectedPoint,
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
                if (IsPointInsideRoomRectangle2D(roomFi, point, tolerance))
                    return true;
            }

            return false;
        }

        private static bool IsPointInsideRoomRectangle2D(FamilyInstance roomFi, XYZ worldPoint, double tolerance)
        {
            if (roomFi == null || worldPoint == null)
                return false;

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                return false;

            Transform inverse;
            try
            {
                inverse = tr.Inverse;
            }
            catch
            {
                return false;
            }

            if (inverse == null)
                return false;

            double width;
            double depth;

            try
            {
                width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
                depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");
            }
            catch
            {
                return false;
            }

            if (width <= 0 || depth <= 0)
                return false;

            XYZ localPoint = inverse.OfPoint(worldPoint);
            return localPoint.X >= -width / 2.0 - tolerance &&
                   localPoint.X <= width / 2.0 + tolerance &&
                   localPoint.Y >= -depth / 2.0 - tolerance &&
                   localPoint.Y <= depth / 2.0 + tolerance;
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

        private static List<Wall> GetExistingWallsFromLineInfo(Document doc, List<ExistingWallLineInfo> existingWallsOnLevel)
        {
            List<Wall> result = new List<Wall>();
            if (doc == null || existingWallsOnLevel == null || existingWallsOnLevel.Count == 0)
                return result;

            HashSet<long> usedIds = new HashSet<long>();

            foreach (ExistingWallLineInfo info in existingWallsOnLevel)
            {
                if (info == null || info.WallId == null || info.WallId == ElementId.InvalidElementId)
                    continue;

                long idValue = IDHelper.ElIdValue(info.WallId);
                if (usedIds.Contains(idValue))
                    continue;

                Wall wall = doc.GetElement(info.WallId) as Wall;
                if (wall == null)
                    continue;

                usedIds.Add(idValue);
                result.Add(wall);
            }

            return result;
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

        private bool TryFindBestHostWallForDoor(XYZ doorPoint, List<Wall> candidateWalls, double maxDistanceToWallAxis, out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            return TryFindBestHostWallForDoor(
                doorPoint,
                candidateWalls,
                maxDistanceToWallAxis,
                false,
                out bestWall,
                out bestProjectedPoint,
                out bestDistance);
        }

        private bool TryFindBestHostWallForDoor(XYZ doorPoint, List<Wall> candidateWalls, double maxDistanceToWallAxis, bool includeWallHalfWidth,
            out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            bestWall = null;
            bestProjectedPoint = null;
            bestDistance = double.MaxValue;
            double bestScore = double.MaxValue;

            if (doorPoint == null || candidateWalls == null || candidateWalls.Count == 0)
                return false;

            foreach (Wall wall in candidateWalls)
            {
                if (wall == null)
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line wallLine = lc.Curve as Line;
                if (wallLine == null || wallLine.Length < 1e-9)
                    continue;

                XYZ projectedPoint;
                double distance;

                if (!TryProjectPointToSegment2D(doorPoint, wallLine, out projectedPoint, out distance))
                    continue;

                double wallHalfWidth = includeWallHalfWidth ? GetWallHalfWidth(wall) : 0;
                double allowedDistance = maxDistanceToWallAxis + wallHalfWidth;

                if (distance > allowedDistance)
                    continue;

                double score = includeWallHalfWidth
                    ? Math.Max(0, distance - wallHalfWidth)
                    : distance;

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

        private static double GetWallHalfWidth(Wall wall)
        {
            if (wall == null)
                return 0;

            try
            {
                return wall.Width / 2.0;
            }
            catch
            {
                return 0;
            }
        }

        private bool TryProjectPointToSegment2D(XYZ point, Line segment, out XYZ projectedPoint, out double distance)
        {
            projectedPoint = null;
            distance = double.MaxValue;

            if (point == null || segment == null)
                return false;

            XYZ a = segment.GetEndPoint(0);
            XYZ b = segment.GetEndPoint(1);
            XYZ ab = b - a;

            double len2 = ab.X * ab.X + ab.Y * ab.Y;
            if (len2 < 1e-12)
                return false;

            double t = ((point.X - a.X) * ab.X + (point.Y - a.Y) * ab.Y) / len2;

            if (t < 0.0)
                t = 0.0;
            else if (t > 1.0)
                t = 1.0;

            double x = a.X + ab.X * t;
            double y = a.Y + ab.Y * t;
            double z = point.Z;

            projectedPoint = new XYZ(x, y, z);
            distance = Distance2D(point, projectedPoint);

            return true;
        }

        private static Level ResolveBaseLevelForPreset(Document doc, ApartmentPresetData preset, ViewPlan targetPlan)
        {
            if (preset != null && !string.IsNullOrWhiteSpace(preset.LowerConstraint))
            {
                string[] parts = preset.LowerConstraint
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToArray();

                foreach (string levelName in parts)
                {
                    Level lvl = FindLevelByName(doc, levelName);
                    if (lvl != null)
                        return lvl;
                }
            }

            if (targetPlan != null && targetPlan.GenLevel != null)
                return targetPlan.GenLevel;

            return null;
        }

        private static Level ResolveTopLevelForPreset(Document doc, ApartmentPresetData preset)
        {
            if (preset == null)
                return null;

            if (string.IsNullOrWhiteSpace(preset.UpperConstraint))
                return null;

            if (string.Equals(preset.UpperConstraint, "Неприсоединённая", StringComparison.OrdinalIgnoreCase))
                return null;

            return FindLevelByName(doc, preset.UpperConstraint.Trim());
        }

        private static Level FindLevelByName(Document doc, string levelName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(levelName))
                return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(x => string.Equals(x.Name, levelName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetSelectedWallTypeNameForThickness(ApartmentPresetData preset, int thicknessMm)
        {
            if (preset == null || preset.WallTypeByThickness == null || preset.WallTypeByThickness.Count == 0)
                return null;

            string value;
            if (preset.WallTypeByThickness.TryGetValue(thicknessMm, out value))
                return value;

            return null;
        }

        private static WallType FindWallTypeByExactSelectionAndThickness(Document doc, string selectedWallTypeName, int thicknessMm)
        {
            if (doc == null || string.IsNullOrWhiteSpace(selectedWallTypeName))
                return null;

            List<WallType> allWallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .ToList();

            WallType exact = allWallTypes.FirstOrDefault(x => string.Equals(x.Name, selectedWallTypeName, StringComparison.OrdinalIgnoreCase));

            if (exact != null)
            {
                int exactThicknessMm;
                if (TryGetWallTypeThicknessMm(exact, out exactThicknessMm) && exactThicknessMm == thicknessMm)
                    return exact;
            }

            foreach (WallType wt in allWallTypes)
            {
                if (wt == null || !string.Equals(wt.Name, selectedWallTypeName, StringComparison.OrdinalIgnoreCase))
                    continue;

                int wallTypeThicknessMm;
                if (!TryGetWallTypeThicknessMm(wt, out wallTypeThicknessMm))
                    continue;

                if (wallTypeThicknessMm == thicknessMm)
                    return wt;
            }

            return null;
        }

        private static WallType FindWallTypeByName(Document doc, string selectedWallTypeName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(selectedWallTypeName))
                return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(x => x != null && string.Equals(x.Name, selectedWallTypeName, StringComparison.OrdinalIgnoreCase));
        }

        private FamilySymbol GetOppositeDoorSymbol(Document doc, FamilySymbol currentSymbol)
        {
            if (doc == null || currentSymbol == null || currentSymbol.Family == null)
                return currentSymbol;

            string currentTypeName = currentSymbol.Name ?? "";
            DoorOpeningMarker currentMarker = GetDoorOpeningMarkerFromTypeName(currentTypeName);

            if (currentMarker == DoorOpeningMarker.None)
                return currentSymbol;

            DoorOpeningMarker oppositeMarker =
                currentMarker == DoorOpeningMarker.Left
                    ? DoorOpeningMarker.Right
                    : DoorOpeningMarker.Left;

            string targetTypeName = ReplaceDoorOpeningMarker(currentTypeName, oppositeMarker);

            FamilySymbol result = FindFamilySymbolByTypeName(doc, currentSymbol.Family, targetTypeName);
            if (result != null)
                return result;

            if (oppositeMarker == DoorOpeningMarker.Right)
            {
                string altRightTypeName = ReplaceDoorOpeningMarker(currentTypeName, DoorOpeningMarker.RightAlt);
                result = FindFamilySymbolByTypeName(doc, currentSymbol.Family, altRightTypeName);
                if (result != null)
                    return result;
            }

            return currentSymbol;
        }

        private static List<ExistingWallLineInfo> GetExistingWallLinesOnLevel(Document doc, ElementId levelId)
        {
            List<ExistingWallLineInfo> result = new List<ExistingWallLineInfo>();
            Level targetLevel = doc != null ? doc.GetElement(levelId) as Level : null;
            double? targetElevation = targetLevel != null ? (double?)targetLevel.Elevation : null;

            IEnumerable<Wall> walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>();

            foreach (Wall wall in walls)
            {
                if (wall == null)
                    continue;

                if (!WallBelongsToOrCrossesLevel(wall, levelId, targetElevation))
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line line = lc.Curve as Line;
                if (line == null || line.Length < 1e-9)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);

                if (targetElevation.HasValue)
                {
                    p0 = new XYZ(p0.X, p0.Y, targetElevation.Value);
                    p1 = new XYZ(p1.X, p1.Y, targetElevation.Value);
                }

                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                double thicknessInternal = 0;
                try
                {
                    thicknessInternal = wall.Width;
                }
                catch
                {
                    thicknessInternal = 0;
                }

                result.Add(new ExistingWallLineInfo
                {
                    WallId = wall.Id,
                    P0 = p0,
                    P1 = p1,
                    Dir = dir,
                    Z = 0.5 * (p0.Z + p1.Z),
                    ThicknessInternal = thicknessInternal
                });
            }

            return result;
        }

        private static bool WallBelongsToOrCrossesLevel(Wall wall, ElementId levelId, double? targetElevation)
        {
            if (wall == null || levelId == null || levelId == ElementId.InvalidElementId)
                return false;

            if (wall.LevelId == levelId)
                return true;

            if (!targetElevation.HasValue)
                return false;

            try
            {
                BoundingBoxXYZ bbox = wall.get_BoundingBox(null);
                if (bbox == null)
                    return false;

                double tol = IDHelper.ConvertMmToInternal(100);
                return bbox.Min.Z <= targetElevation.Value + tol &&
                       bbox.Max.Z >= targetElevation.Value - tol;
            }
            catch
            {
                return false;
            }
        }

        private static void ApplyWallPresetParameters(Wall wall, Level baseLevel, Level topLevel, double baseOffsetInternal, double unconnectedHeightInternal)
        {
            if (wall == null)
                return;

            Parameter pBaseConstraint = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
            if (pBaseConstraint != null && !pBaseConstraint.IsReadOnly && baseLevel != null)
                pBaseConstraint.Set(baseLevel.Id);

            Parameter pBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
            if (pBaseOffset != null && !pBaseOffset.IsReadOnly)
                pBaseOffset.Set(baseOffsetInternal);

            Parameter pTopConstraint = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            Parameter pUnconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);

            if (topLevel != null)
            {
                if (pTopConstraint != null && !pTopConstraint.IsReadOnly)
                    pTopConstraint.Set(topLevel.Id);
            }
            else
            {
                if (pUnconnectedHeight != null && !pUnconnectedHeight.IsReadOnly)
                    pUnconnectedHeight.Set(unconnectedHeightInternal);
            }
        }

        private static Line BuildShaftWallAxisFromInteriorFaceLine(Line faceLine, WallType wallType, XYZ apartmentInteriorPoint)
        {
            if (faceLine == null)
                return null;

            XYZ p0 = faceLine.GetEndPoint(0);
            XYZ p1 = faceLine.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);
            if (dir == null)
                return faceLine;

            double wallWidth = GetWallTypeWidthInternal(wallType);
            if (wallWidth <= 1e-9)
                return faceLine;

            XYZ normal = new XYZ(-dir.Y, dir.X, 0);
            XYZ mid = new XYZ(
                0.5 * (p0.X + p1.X),
                0.5 * (p0.Y + p1.Y),
                0.5 * (p0.Z + p1.Z));

            XYZ offsetDir = normal;
            if (apartmentInteriorPoint != null)
            {
                XYZ toInterior = Normalize2D(apartmentInteriorPoint - mid);
                if (toInterior != null && Dot2D(offsetDir, toInterior) > 0)
                    offsetDir = new XYZ(-offsetDir.X, -offsetDir.Y, 0);
            }

            double halfWidth = wallWidth / 2.0;
            XYZ offset = new XYZ(offsetDir.X * halfWidth, offsetDir.Y * halfWidth, 0);

            return Line.CreateBound(
                new XYZ(p0.X + offset.X, p0.Y + offset.Y, p0.Z),
                new XYZ(p1.X + offset.X, p1.Y + offset.Y, p1.Z));
        }

        private static double GetWallTypeWidthInternal(WallType wallType)
        {
            if (wallType == null)
                return 0;

            try
            {
                if (wallType.Width > 0)
                    return wallType.Width;
            }
            catch
            {
            }

            int thicknessMm;
            if (TryGetWallTypeThicknessMm(wallType, out thicknessMm) && thicknessMm > 0)
                return IDHelper.ConvertMmToInternal(thicknessMm);

            return 0;
        }

        private static List<ExistingWallLineInfo> BuildWallLineInfoFromWalls(IEnumerable<Wall> walls)
        {
            List<ExistingWallLineInfo> result = new List<ExistingWallLineInfo>();
            if (walls == null)
                return result;

            foreach (Wall wall in walls)
            {
                if (wall == null || !wall.IsValidObject)
                    continue;

                Line line = GetWallAxisLine(wall);
                if (line == null || line.Length < 1e-9)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                result.Add(new ExistingWallLineInfo
                {
                    WallId = wall.Id,
                    P0 = p0,
                    P1 = p1,
                    Dir = dir,
                    Z = 0.5 * (p0.Z + p1.Z),
                    ThicknessInternal = GetWallHalfWidth(wall) * 2.0
                });
            }

            return result;
        }

        private PreparedApartmentDoors PrepareDoorsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, List<string> debugMessages, ApartmentProcessState state = null)
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

                XYZ insertPointInProject = doorTransform.Origin;
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

            AddApartmentFamilyDoorMarkersFallback(doc, apartmentFi, preset, debugMessages, state, result);

            return result;
        }

        private void AddApartmentFamilyDoorMarkersFallback(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset,
            List<string> debugMessages, ApartmentProcessState state, PreparedApartmentDoors result)
        {
            if (doc == null || apartmentFi == null || result == null || apartmentFi.Symbol == null || apartmentFi.Symbol.Family == null)
                return;

            List<FamilyRoomMarker> rooms = new List<FamilyRoomMarker>();
            List<FamilyDoorMarker> doors = new List<FamilyDoorMarker>();
            HashSet<string> visitedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            CollectApartmentFamilyMarkersRecursive(
                doc,
                apartmentFi.Symbol.Family,
                Transform.Identity,
                rooms,
                doors,
                null,
                visitedFamilies);

            if (doors.Count == 0)
                return;

            Transform apartmentTransform = apartmentFi.GetTransform();
            if (apartmentTransform == null)
                apartmentTransform = Transform.Identity;

            foreach (FamilyDoorMarker door in doors)
            {
                if (door == null || string.IsNullOrWhiteSpace(door.DoorTypeName2D) || door.DoorWidthMm <= 0 || door.LocalPoint == null)
                    continue;

                XYZ insertPointInProject = apartmentTransform.OfPoint(door.LocalPoint);
                PreparedDoorPlacement duplicateDoor = FindDuplicatePreparedDoor(result, door.DoorWidthMm, insertPointInProject);
                if (duplicateDoor != null)
                {
                    if (!door.IsEntranceDoor || duplicateDoor.IsEntranceDoor)
                        continue;

                    result.Doors.Remove(duplicateDoor);
                }

                XYZ sourceHandDirection = null;
                XYZ sourceFacingDirection = null;

                if (door.LocalTransform != null)
                {
                    sourceHandDirection = Normalize2D(apartmentTransform.OfVector(door.LocalTransform.BasisX));
                    sourceFacingDirection = Normalize2D(apartmentTransform.OfVector(door.LocalTransform.BasisY));
                }

                TryAddPreparedDoorPlacement(
                    doc,
                    apartmentFi,
                    preset,
                    debugMessages,
                    state,
                    result,
                    null,
                    ElementId.InvalidElementId,
                    door.DoorTypeName2D,
                    door.Comment,
                    door.IsEntranceDoor,
                    door.DoorWidthMm,
                    insertPointInProject,
                    sourceHandDirection,
                    sourceFacingDirection,
                    true);
            }
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
                preset.DoorsByRoomCategory.TryGetValue(presetKey, out selectedDoorTypeName);

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

            if (isEntranceDoor)
            {
                try
                {
                    DoorTypeMirrorEnsureResult ensureResult = EnsureDoorMirrorTypeExists(doc, baseDoorSymbol);
                    if (ensureResult != null && ensureResult.HasMessage)
                        entranceDiagnostics.Add(ensureResult.Message);
                }
                catch (Exception ex)
                {
                    AddApartmentDiagnostic(
                        state,
                        debugMessages,
                        "Не удалось подготовить парный тип входной двери для '" +
                        BuildDoorTypeDisplayName(baseDoorSymbol) + "': " + ex.Message);
                }
            }

            FamilySymbol resolvedDoorSymbol = isEntranceDoor || sourceDoorFi == null
                ? baseDoorSymbol
                : ResolveDoorSymbolForPlacement(doc, sourceDoorFi, baseDoorSymbol, debugMessages);

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
                IsEntranceDoor = isEntranceDoor,
                Diagnostics = entranceDiagnostics,
                RequiresOppositeDoorTypeAfterWallFlip = false
            });

            return true;
        }

        private static PreparedDoorPlacement FindDuplicatePreparedDoor(PreparedApartmentDoors result, int widthMm, XYZ insertPoint)
        {
            if (result == null || result.Doors == null || insertPoint == null)
                return null;

            double pointTolerance = IDHelper.ConvertMmToInternal(10);

            return result.Doors.FirstOrDefault(x =>
                x != null &&
                x.DoorWidthMm == widthMm &&
                x.InsertPoint != null &&
                Distance2D(x.InsertPoint, insertPoint) <= pointTolerance);
        }

        private static string FormatElementIdForDiagnostic(ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId)
                return "<из семейства>";

            return IDHelper.ElIdValue(id).ToString();
        }

        private PreparedApartmentWindows PrepareWindowsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset,
            List<string> debugMessages, ApartmentProcessState state = null)
        {
            PreparedApartmentWindows result = new PreparedApartmentWindows();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyWindowMarker> windowMarkers = CollectWindowMarkersFromApartmentInstance(doc, apartmentFi);
            if (windowMarkers == null || windowMarkers.Count == 0)
                return result;

            string selectedWindowType = preset != null ? preset.WindowType : null;
            if (string.IsNullOrWhiteSpace(selectedWindowType) ||
                string.Equals(selectedWindowType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Для окон квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) + " не выбран тип окна.");
                return result;
            }

            FamilySymbol windowSymbol = FindWindowSymbolByDisplayName(doc, selectedWindowType);
            if (windowSymbol == null)
            {
                AddApartmentDiagnostic(
                    state,
                    debugMessages,
                    "Для окон квартиры ID = " + IDHelper.ElIdValue(apartmentFi.Id) +
                    " выбран тип '" + selectedWindowType + "', но такой тип не найден в проекте.");
                return result;
            }

            double sillHeightInternal = IDHelper.ConvertMmToInternal(preset != null && preset.WindowSillHeight > 0 ? preset.WindowSillHeight : 900);

            foreach (FamilyWindowMarker marker in windowMarkers)
            {
                if (marker == null || marker.LocalP0 == null || marker.LocalP1 == null)
                    continue;

                XYZ p0 = marker.LocalP0;
                XYZ p1 = marker.LocalP1;

                if (p0 == null || p1 == null || Distance2D(p0, p1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                Line sourceLine = Line.CreateBound(p0, p1);
                XYZ insertPoint = new XYZ(
                    0.5 * (p0.X + p1.X),
                    0.5 * (p0.Y + p1.Y),
                    0.5 * (p0.Z + p1.Z));

                result.Windows.Add(new PreparedWindowPlacement
                {
                    ApartmentId = apartmentFi.Id,
                    WindowSymbol = windowSymbol,
                    SourceLine = sourceLine,
                    InsertPoint = insertPoint,
                    ReferenceDirection = Normalize2D(p1 - p0),
                    SillHeightInternal = sillHeightInternal,
                    Diagnostics = new List<string>()
                });
            }

            return result;
        }

        private static XYZ GetApartmentInteriorReferencePoint(FamilyInstance apartmentFi, Document doc)
        {
            List<FamilyInstance> rooms = FindRoomSubComponents(doc, apartmentFi);
            if (rooms != null && rooms.Count > 0)
            {
                double x = 0;
                double y = 0;
                double z = 0;
                int count = 0;

                foreach (FamilyInstance roomFi in rooms)
                {
                    XYZ center = GetRoomCenterPoint(roomFi);
                    if (center == null)
                        continue;

                    x += center.X;
                    y += center.Y;
                    z += center.Z;
                    count++;
                }

                if (count > 0)
                    return new XYZ(x / count, y / count, z / count);
            }

            Transform tr = apartmentFi != null ? apartmentFi.GetTransform() : null;
            return tr != null ? tr.Origin : null;
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

            using (Transaction t = new Transaction(doc, "KPLN. Создание парных типов двери"))
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

        private FamilySymbol ResolveDoorSymbolForPlacement(Document doc, FamilyInstance source2DDoor, FamilySymbol baseDoorSymbol, List<string> debugMessages)
        {
            if (doc == null || source2DDoor == null || baseDoorSymbol == null)
                return baseDoorSymbol;

            string baseTypeName = baseDoorSymbol.Name ?? "";
            DoorOpeningMarker baseMarker = GetDoorOpeningMarkerFromTypeName(baseTypeName);

            if (baseMarker == DoorOpeningMarker.None)
                return baseDoorSymbol;

            DoorTypeMirrorEnsureResult ensureResult = EnsureDoorMirrorTypeExists(doc, baseDoorSymbol);
            if (ensureResult != null && ensureResult.HasMessage && debugMessages != null)
                debugMessages.Add(ensureResult.Message);

            bool faceFlip;
            bool mirrored;
            if (!TryGetDoorOrientationFlags(source2DDoor, out faceFlip, out mirrored))
                return baseDoorSymbol;

            DoorOpeningMarker desiredMarker = ResolveDesiredDoorOpeningMarker(faceFlip, mirrored);
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

        private static bool TryGetDoorOrientationFlags(FamilyInstance doorFi, out bool faceFlip, out bool mirrored)
        {
            faceFlip = false;
            mirrored = false;

            if (doorFi == null)
                return false;

            bool hasFaceFlip = TryGetYesNoParamFromElementOrType(doorFi, out faceFlip, "faceFlip", "FaceFlip", "Facing Flip", "FacingFlipped");
            bool hasMirrored = TryGetYesNoParamFromElementOrType(doorFi, out mirrored, "mirrored", "Mirrored");

            if (!hasFaceFlip)
            {
                try
                {
                    faceFlip = doorFi.FacingFlipped;
                    hasFaceFlip = true;
                }
                catch
                {
                }
            }

            if (!hasMirrored)
            {
                try
                {
                    mirrored = doorFi.Mirrored;
                    hasMirrored = true;
                }
                catch
                {
                }
            }

            return hasFaceFlip && hasMirrored;
        }

        private static bool TryGetYesNoParamFromElementOrType(Element e, out bool value, params string[] paramNames)
        {
            value = false;

            if (e == null || paramNames == null || paramNames.Length == 0)
                return false;

            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null)
                {
                    if (p.StorageType == StorageType.Integer)
                    {
                        value = p.AsInteger() != 0;
                        return true;
                    }

                    if (p.StorageType == StorageType.String)
                    {
                        string s = p.AsString();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            s = s.Trim();
                            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                            {
                                value = true;
                                return true;
                            }

                            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                            {
                                value = false;
                                return true;
                            }
                        }
                    }
                }
            }

            Element typeElem = null;
            if (e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null)
                    {
                        if (p.StorageType == StorageType.Integer)
                        {
                            value = p.AsInteger() != 0;
                            return true;
                        }

                        if (p.StorageType == StorageType.String)
                        {
                            string s = p.AsString();
                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                s = s.Trim();
                                if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "1", StringComparison.OrdinalIgnoreCase))
                                {
                                    value = true;
                                    return true;
                                }

                                if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(s, "0", StringComparison.OrdinalIgnoreCase))
                                {
                                    value = false;
                                    return true;
                                }
                            }
                        }
                    }
                }
            }

            return false;
        }

        private static DoorOpeningMarker ResolveDesiredDoorOpeningMarker(bool faceFlip, bool mirrored)
        {
            if (!faceFlip && !mirrored)
                return DoorOpeningMarker.Left;

            if (faceFlip && mirrored)
                return DoorOpeningMarker.Left;

            if (faceFlip && !mirrored)
                return DoorOpeningMarker.Right;

            if (!faceFlip && mirrored)
                return DoorOpeningMarker.Right;

            return DoorOpeningMarker.None;
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
                    result.Add(subFi);

                CollectDoorSubComponentsRecursive(doc, subFi, result);
            }
        }

        private PreparedApartmentRooms PrepareRoomsForApartment(Document doc, FamilyInstance apartmentFi, List<string> debugMessages)
        {
            PreparedApartmentRooms result = new PreparedApartmentRooms();
            result.ApartmentId = apartmentFi != null ? apartmentFi.Id : ElementId.InvalidElementId;

            if (doc == null || apartmentFi == null)
                return result;

            List<FamilyInstance> roomInstances = FindRoomSubComponents(doc, apartmentFi);
            if (roomInstances == null || roomInstances.Count == 0)
                return result;

            foreach (FamilyInstance roomFi in roomInstances)
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

                    Transform roomTransform = roomFi.GetTransform();
                    if (roomTransform == null)
                        continue;

                    XYZ insertPointInProject = roomTransform.Origin;

                    result.Rooms.Add(new PreparedRoomPlacement
                    {
                        ApartmentId = apartmentFi.Id,
                        RoomName = roomName.Trim(),
                        InsertPoint = insertPointInProject,
                        ExpectedAreaInternal = expectedAreaInternal
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

            return result;
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

        private int PlaceRoomsForApartment(Document doc, PreparedApartmentRooms apartmentRooms, Level roomLevel,
            List<RoomAreaMismatchInfo> roomAreaMismatches, List<DeletedRoomMismatchInfo> deletedRoomMismatches, List<Wall> createdWallsForApartment,
            List<ElementId> createdRoomIds = null)
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

                    if (HasRoomAtPoint(doc, roomLevel, roomPoint))
                        continue;

                    UV roomUv = new UV(preparedRoom.InsertPoint.X, preparedRoom.InsertPoint.Y);

                    Room createdRoom = doc.Create.NewRoom(roomLevel, roomUv);
                    if (createdRoom == null)
                        continue;

                    Parameter roomNameParam = createdRoom.get_Parameter(BuiltInParameter.ROOM_NAME);
                    if (roomNameParam != null && !roomNameParam.IsReadOnly && !string.IsNullOrWhiteSpace(preparedRoom.RoomName))
                        roomNameParam.Set(preparedRoom.RoomName);

                    if (preparedRoom.ExpectedAreaInternal > 0)
                    {
                        double actualAreaInternal = createdRoom.Area;
                        double expectedAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(preparedRoom.ExpectedAreaInternal);
                        double actualAreaSquareMeters = IDHelper.ConvertInternalAreaToSquareMeters(actualAreaInternal);

                        if (Math.Abs(expectedAreaSquareMeters - actualAreaSquareMeters) > areaToleranceSquareMeters)
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

                            if (deletedRoomMismatches != null)
                            {
                                deletedRoomMismatches.Add(new DeletedRoomMismatchInfo
                                {
                                    ApartmentId = preparedRoom.ApartmentId,
                                    RoomName = preparedRoom.RoomName,
                                    ExpectedAreaInternal = preparedRoom.ExpectedAreaInternal,
                                    ActualAreaInternal = actualAreaInternal
                                });
                            }

                            doc.Delete(createdRoom.Id);
                            DeleteIsolatedCreatedWalls(doc, createdWallsForApartment);
                            continue;
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

        #region ГЕОМЕТРИЯ ПОМЕЩЕНИЙ И ПОСТРОЕНИЕ ОСЕЙ

        private static CurveLoop BuildRoomLoopFromInstance(FamilyInstance roomFi)
        {
            if (roomFi == null)
                throw new ArgumentNullException("roomFi");

            double width = GetRequiredLengthParam(roomFi, "Ширина", "Width");
            double depth = GetRequiredLengthParam(roomFi, "Глубина", "Depth");

            Transform tr = roomFi.GetTransform();
            if (tr == null)
                throw new Exception("Не удалось получить Transform для вложенного помещения.");

            double halfW = width / 2.0;
            double halfD = depth / 2.0;

            XYZ p1 = tr.OfPoint(new XYZ(-halfW, -halfD, 0));
            XYZ p2 = tr.OfPoint(new XYZ(halfW, -halfD, 0));
            XYZ p3 = tr.OfPoint(new XYZ(halfW, halfD, 0));
            XYZ p4 = tr.OfPoint(new XYZ(-halfW, halfD, 0));

            List<Curve> profile = new List<Curve>();
            profile.Add(Line.CreateBound(p1, p2));
            profile.Add(Line.CreateBound(p2, p3));
            profile.Add(Line.CreateBound(p3, p4));
            profile.Add(Line.CreateBound(p4, p1));

            return CurveLoop.Create(profile);
        }

        private static double GetRequiredLengthParam(Element e, params string[] paramNames)
        {
            foreach (string paramName in paramNames)
            {
                Parameter p = e.LookupParameter(paramName);
                if (p != null && p.StorageType == StorageType.Double)
                    return p.AsDouble();
            }

            Element typeElem = null;
            if (e != null && e.Document != null)
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElem = e.Document.GetElement(typeId);
            }

            if (typeElem != null)
            {
                foreach (string paramName in paramNames)
                {
                    Parameter p = typeElem.LookupParameter(paramName);
                    if (p != null && p.StorageType == StorageType.Double)
                        return p.AsDouble();
                }
            }

            throw new Exception("Не найден параметр: " + string.Join(", ", paramNames));
        }

        private static List<Line> BuildClosedWallAxisLinesFromRooms(List<CurveLoop> roomLoops, double wallWidth, List<string> debugMessages)
        {
            List<Line> result = new List<Line>();
            double halfWidth = wallWidth / 2.0;

            foreach (CurveLoop roomLoop in roomLoops)
            {
                try
                {
                    List<Line> offsetLoopLines = BuildOffsetClosedLoop(roomLoop, halfWidth);
                    result.AddRange(offsetLoopLines);
                }
                catch (Exception ex)
                {
                    debugMessages.Add("Ошибка построения замкнутого контура стены: " + ex.Message);
                }
            }

            return MergeCollinearLines(result);
        }

        private static List<Line> BuildOffsetClosedLoop(CurveLoop loop, double offset)
        {
            const double tol = 1e-9;

            List<Line> edges = loop.Cast<Curve>()
                .Select(x => x as Line)
                .Where(x => x != null && x.Length > tol)
                .ToList();

            if (edges.Count < 3)
                throw new Exception("Контур помещения должен содержать минимум 3 линейных сегмента.");

            List<XYZ> vertices = ExtractOrderedVertices(edges);
            if (vertices.Count < 3)
                throw new Exception("Не удалось извлечь вершины контура помещения.");

            bool ccw = GetSignedAreaXY(vertices) > 0.0;

            List<OffsetLine2D> offsetLines = new List<OffsetLine2D>();

            for (int i = 0; i < vertices.Count; i++)
            {
                XYZ a = vertices[i];
                XYZ b = vertices[(i + 1) % vertices.Count];
                XYZ dir = Normalize2D(b - a);

                if (dir == null)
                    throw new Exception("Обнаружен нулевой сегмент в контуре.");

                XYZ outward = ccw ? new XYZ(dir.Y, -dir.X, 0) : new XYZ(-dir.Y, dir.X, 0);

                XYZ oa = new XYZ(a.X + outward.X * offset, a.Y + outward.Y * offset, a.Z);
                XYZ ob = new XYZ(b.X + outward.X * offset, b.Y + outward.Y * offset, b.Z);

                offsetLines.Add(new OffsetLine2D
                {
                    P0 = oa,
                    P1 = ob,
                    Dir = dir,
                    Z = a.Z
                });
            }

            List<XYZ> offsetVertices = new List<XYZ>();

            for (int i = 0; i < offsetLines.Count; i++)
            {
                OffsetLine2D prev = offsetLines[(i - 1 + offsetLines.Count) % offsetLines.Count];
                OffsetLine2D cur = offsetLines[i];

                XYZ intersection;
                if (!TryIntersectInfiniteLines2D(prev.P0, prev.Dir, cur.P0, cur.Dir, out intersection))
                    intersection = cur.P0;

                double z = 0.5 * (prev.Z + cur.Z);
                offsetVertices.Add(new XYZ(intersection.X, intersection.Y, z));
            }

            List<Line> result = new List<Line>();

            for (int i = 0; i < offsetVertices.Count; i++)
            {
                XYZ p0 = offsetVertices[i];
                XYZ p1 = offsetVertices[(i + 1) % offsetVertices.Count];

                if (p0.DistanceTo(p1) > tol)
                    result.Add(Line.CreateBound(p0, p1));
            }

            return result;
        }

        private static List<XYZ> ExtractOrderedVertices(List<Line> lines)
        {
            const double tol = 1e-6;
            List<XYZ> pts = new List<XYZ>();

            if (lines == null || lines.Count == 0)
                return pts;

            foreach (Line line in lines)
            {
                XYZ p = line.GetEndPoint(0);
                if (pts.Count == 0 || pts[pts.Count - 1].DistanceTo(p) > tol)
                    pts.Add(p);
            }

            if (pts.Count >= 2 && pts[0].DistanceTo(pts[pts.Count - 1]) < tol)
                pts.RemoveAt(pts.Count - 1);

            return pts;
        }

        private static double GetSignedAreaXY(List<XYZ> pts)
        {
            if (pts == null || pts.Count < 3)
                return 0.0;

            double area2 = 0.0;

            for (int i = 0; i < pts.Count; i++)
            {
                XYZ a = pts[i];
                XYZ b = pts[(i + 1) % pts.Count];
                area2 += (a.X * b.Y) - (b.X * a.Y);
            }

            return area2 * 0.5;
        }

        private static List<Line> MergeCollinearLines(List<Line> lines)
        {
            const double tol = 1e-6;
            List<GenericAxisLineData> data = new List<GenericAxisLineData>();

            foreach (Line line in lines)
            {
                if (line == null || line.Length <= tol)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ dir = Normalize2D(p1 - p0);
                if (dir == null)
                    continue;

                dir = CanonicalizeDirection(dir);
                XYZ normal = new XYZ(-dir.Y, dir.X, 0);
                double offset = Dot2D(p0, normal);
                double z = 0.5 * (p0.Z + p1.Z);
                double t0 = Dot2D(p0, dir);
                double t1 = Dot2D(p1, dir);
                double from = Math.Min(t0, t1);
                double to = Math.Max(t0, t1);

                data.Add(new GenericAxisLineData
                {
                    Dir = dir,
                    Normal = normal,
                    Offset = offset,
                    Z = z,
                    From = from,
                    To = to
                });
            }

            List<IGrouping<GenericAxisGroupKey, GenericAxisLineData>> groups = data
                .GroupBy(x => new GenericAxisGroupKey(x.Dir, x.Offset, x.Z, tol))
                .ToList();

            List<Line> result = new List<Line>();

            foreach (IGrouping<GenericAxisGroupKey, GenericAxisLineData> group in groups)
            {
                List<GenericAxisLineData> ordered = group.OrderBy(x => x.From).ThenBy(x => x.To).ToList();

                if (ordered.Count == 0)
                    continue;

                double curFrom = ordered[0].From;
                double curTo = ordered[0].To;

                for (int i = 1; i < ordered.Count; i++)
                {
                    GenericAxisLineData next = ordered[i];

                    if (next.From <= curTo + tol)
                    {
                        curTo = Math.Max(curTo, next.To);
                    }
                    else
                    {
                        result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
                        curFrom = next.From;
                        curTo = next.To;
                    }
                }

                result.Add(BuildGenericAxisLine(group.Key.Dir, group.Key.Offset, group.Key.Z, curFrom, curTo));
            }

            return result;
        }

        private static Line BuildGenericAxisLine(XYZ dir, double offset, double z, double from, double to)
        {
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            XYZ p0 = new XYZ(dir.X * from + normal.X * offset, dir.Y * from + normal.Y * offset, z);
            XYZ p1 = new XYZ(dir.X * to + normal.X * offset, dir.Y * to + normal.Y * offset, z);

            return Line.CreateBound(p0, p1);
        }

        private static List<Line> SnapNewLinesToExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            List<Line> result = new List<Line>();

            foreach (Line line in newLines)
            {
                if (line == null || line.Length < 1e-9)
                    continue;

                Line snapped = SnapSingleLineToExistingWalls(line, existingWalls, snapTol);
                if (snapped != null && snapped.Length > 1e-9)
                    result.Add(snapped);
            }

            return result;
        }

        private static Line SnapSingleLineToExistingWalls(Line line, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);

            if (dir == null)
                return line;

            XYZ newP0 = SnapEndpointToExistingWalls(p0, new XYZ(-dir.X, -dir.Y, 0), line, existingWalls, snapTol);
            XYZ newP1 = SnapEndpointToExistingWalls(p1, dir, line, existingWalls, snapTol);

            if (newP0.DistanceTo(newP1) < 1e-9)
                return null;

            return Line.CreateBound(newP0, newP1);
        }

        private static XYZ SnapEndpointToExistingWalls(XYZ endpoint, XYZ extensionDir, Line sourceLine, List<ExistingWallLineInfo> existingWalls, double snapTol)
        {
            const double tol = 1e-9;

            XYZ bestPoint = endpoint;
            double bestDist = double.MaxValue;

            XYZ sourceP0 = sourceLine.GetEndPoint(0);
            XYZ sourceP1 = sourceLine.GetEndPoint(1);
            XYZ sourceDir = Normalize2D(sourceP1 - sourceP0);

            if (sourceDir == null)
                return endpoint;

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (Math.Abs(ex.Z - endpoint.Z) > snapTol)
                    continue;

                XYZ inter;
                if (TryIntersectInfiniteLines2D(endpoint, extensionDir, ex.P0, ex.Dir, out inter))
                {
                    if (PointOnSegment2D(inter, ex.P0, ex.P1, snapTol))
                    {
                        XYZ delta = inter - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, inter);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(inter.X, inter.Y, endpoint.Z);
                        }
                    }
                }

                if (AreCollinear2D(sourceP0, sourceP1, ex.P0, ex.P1, snapTol))
                {
                    XYZ[] candidates = new XYZ[] { ex.P0, ex.P1 };

                    foreach (XYZ c in candidates)
                    {
                        XYZ delta = c - endpoint;
                        double along = Dot2D(delta, extensionDir);
                        double dist = Distance2D(endpoint, c);

                        if (along >= -tol && dist <= snapTol && dist < bestDist)
                        {
                            bestDist = dist;
                            bestPoint = new XYZ(c.X, c.Y, endpoint.Z);
                        }
                    }
                }
            }

            return bestPoint;
        }

        private static List<Line> RemoveSegmentsOverlappingExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls, double candidateThicknessInternal)
        {
            const double tol = 1e-6;
            List<Line> result = new List<Line>();

            foreach (Line newLine in newLines)
            {
                if (newLine == null || newLine.Length <= tol)
                    continue;

                List<Line> remaining = SubtractExistingWallsFromNewLine(newLine, existingWalls, candidateThicknessInternal);

                foreach (Line part in remaining)
                {
                    if (part != null && part.Length > tol)
                        result.Add(part);
                }
            }

            return result;
        }

        private static List<Line> SubtractExistingWallsFromNewLine(Line newLine, List<ExistingWallLineInfo> existingWalls, double candidateThicknessInternal)
        {
            const double tol = 1e-6;

            XYZ p0 = newLine.GetEndPoint(0);
            XYZ p1 = newLine.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);

            if (dir == null)
                return new List<Line>();

            dir = CanonicalizeDirection(dir);
            XYZ normal = new XYZ(-dir.Y, dir.X, 0);

            double newOffset = Dot2D(p0, normal);
            double newZ = 0.5 * (p0.Z + p1.Z);

            double newT0 = Dot2D(p0, dir);
            double newT1 = Dot2D(p1, dir);
            double newFrom = Math.Min(newT0, newT1);
            double newTo = Math.Max(newT0, newT1);

            List<Interval1D> cutIntervals = new List<Interval1D>();

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (ex == null)
                    continue;

                if (Math.Abs(ex.Z - newZ) > tol)
                    continue;

                XYZ exDir = CanonicalizeDirection(ex.Dir);
                if (exDir == null)
                    continue;

                if (Math.Abs(Cross2D(dir, exDir)) > tol)
                    continue;

                double exOffset = Dot2D(ex.P0, normal);
                double existingThicknessInternal = ex.ThicknessInternal > 0
                    ? ex.ThicknessInternal
                    : 0;

                double candidateThickness = candidateThicknessInternal > 0
                    ? candidateThicknessInternal
                    : 0;

                double halfSumThickness = (candidateThickness + existingThicknessInternal) / 2.0;
                if (Math.Abs(exOffset - newOffset) > halfSumThickness + tol)
                    continue;

                double exT0 = Dot2D(ex.P0, dir);
                double exT1 = Dot2D(ex.P1, dir);
                double exFrom = Math.Min(exT0, exT1);
                double exTo = Math.Max(exT0, exT1);

                double overlapFrom = Math.Max(newFrom, exFrom);
                double overlapTo = Math.Min(newTo, exTo);

                if (overlapTo - overlapFrom > tol)
                {
                    cutIntervals.Add(new Interval1D
                    {
                        From = overlapFrom,
                        To = overlapTo
                    });
                }
            }

            if (cutIntervals.Count == 0)
                return new List<Line> { newLine };

            List<Interval1D> mergedCuts = MergeIntervals(cutIntervals);
            List<Interval1D> remainingIntervals = SubtractIntervals(
                new Interval1D { From = newFrom, To = newTo },
                mergedCuts);

            List<Line> result = new List<Line>();

            foreach (Interval1D interval in remainingIntervals)
            {
                if (interval.To - interval.From <= tol)
                    continue;

                XYZ rp0 = new XYZ(dir.X * interval.From + normal.X * newOffset, dir.Y * interval.From + normal.Y * newOffset, newZ);
                XYZ rp1 = new XYZ(dir.X * interval.To + normal.X * newOffset, dir.Y * interval.To + normal.Y * newOffset, newZ);

                if (rp0.DistanceTo(rp1) > tol)
                    result.Add(Line.CreateBound(rp0, rp1));
            }

            return result;
        }

        private static List<Interval1D> MergeIntervals(List<Interval1D> intervals)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (intervals == null || intervals.Count == 0)
                return result;

            List<Interval1D> ordered = intervals
                .Where(x => x != null && x.To - x.From > tol)
                .OrderBy(x => x.From)
                .ThenBy(x => x.To)
                .ToList();

            if (ordered.Count == 0)
                return result;

            double curFrom = ordered[0].From;
            double curTo = ordered[0].To;

            for (int i = 1; i < ordered.Count; i++)
            {
                Interval1D next = ordered[i];

                if (next.From <= curTo + tol)
                {
                    curTo = Math.Max(curTo, next.To);
                }
                else
                {
                    result.Add(new Interval1D { From = curFrom, To = curTo });
                    curFrom = next.From;
                    curTo = next.To;
                }
            }

            result.Add(new Interval1D { From = curFrom, To = curTo });
            return result;
        }

        private static List<Interval1D> SubtractIntervals(Interval1D source, List<Interval1D> cuts)
        {
            const double tol = 1e-6;
            List<Interval1D> result = new List<Interval1D>();

            if (source == null || source.To - source.From <= tol)
                return result;

            if (cuts == null || cuts.Count == 0)
            {
                result.Add(source);
                return result;
            }

            double cursor = source.From;

            foreach (Interval1D cut in cuts)
            {
                if (cut == null || cut.To - cut.From <= tol)
                    continue;

                if (cut.To <= source.From + tol)
                    continue;

                if (cut.From >= source.To - tol)
                    break;

                double cutFrom = Math.Max(source.From, cut.From);
                double cutTo = Math.Min(source.To, cut.To);

                if (cutFrom > cursor + tol)
                {
                    result.Add(new Interval1D
                    {
                        From = cursor,
                        To = cutFrom
                    });
                }

                cursor = Math.Max(cursor, cutTo);
            }

            if (cursor < source.To - tol)
            {
                result.Add(new Interval1D
                {
                    From = cursor,
                    To = source.To
                });
            }

            return result;
        }

        private static bool IntersectsExistingWalls(Line candidate, double candidateThicknessInternal, List<ExistingWallLineInfo> existingWalls, double tol)
        {
            if (candidate == null || existingWalls == null || existingWalls.Count == 0)
                return false;

            XYZ a0 = candidate.GetEndPoint(0);
            XYZ a1 = candidate.GetEndPoint(1);
            XYZ aDir = Normalize2D(a1 - a0);

            if (aDir == null)
                return false;

            double aZ = 0.5 * (a0.Z + a1.Z);

            foreach (ExistingWallLineInfo ex in existingWalls)
            {
                if (ex == null)
                    continue;

                if (Math.Abs(ex.Z - aZ) > tol)
                    continue;

                if (HasForbiddenAxisIntersection2D(a0, a1, ex.P0, ex.P1, tol))
                    return true;

                if (HasParallelWallBodyOverlap(candidate, candidateThicknessInternal, ex, tol))
                    return true;
            }

            return false;
        }

        private static bool HasForbiddenAxisIntersection2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            if (a0 == null || a1 == null || b0 == null || b1 == null)
                return false;

            if (AreCollinear2D(a0, a1, b0, b1, tol))
            {
                XYZ dir = Normalize2D(a1 - a0);
                if (dir == null)
                    return false;

                double aT0 = Dot2D(a0, dir);
                double aT1 = Dot2D(a1, dir);
                double bT0 = Dot2D(b0, dir);
                double bT1 = Dot2D(b1, dir);

                double aFrom = Math.Min(aT0, aT1);
                double aTo = Math.Max(aT0, aT1);
                double bFrom = Math.Min(bT0, bT1);
                double bTo = Math.Max(bT0, bT1);

                double overlapFrom = Math.Max(aFrom, bFrom);
                double overlapTo = Math.Min(aTo, bTo);
                double overlapLen = overlapTo - overlapFrom;

                return overlapLen > tol * 10.0;
            }

            XYZ intersection;
            if (!TryIntersectSegments2D(a0, a1, b0, b1, out intersection, tol))
                return false;

            bool isEndpointOfA =
                Distance2D(intersection, a0) <= tol ||
                Distance2D(intersection, a1) <= tol;

            bool isEndpointOfB =
                Distance2D(intersection, b0) <= tol ||
                Distance2D(intersection, b1) <= tol;

            if (isEndpointOfA || isEndpointOfB)
                return false;

            return true;
        }

        private static bool HasParallelWallBodyOverlap(Line candidate, double candidateThicknessInternal, ExistingWallLineInfo existingWall, double tol)
        {
            if (candidate == null || existingWall == null)
                return false;

            XYZ a0 = candidate.GetEndPoint(0);
            XYZ a1 = candidate.GetEndPoint(1);

            XYZ aDir = Normalize2D(a1 - a0);
            XYZ bDir = CanonicalizeDirection(existingWall.Dir);

            if (aDir == null || bDir == null)
                return false;

            aDir = CanonicalizeDirection(aDir);

            if (Math.Abs(Cross2D(aDir, bDir)) > tol)
                return false;

            XYZ normal = new XYZ(-aDir.Y, aDir.X, 0);

            double aOffset = Dot2D(a0, normal);
            double bOffset = Dot2D(existingWall.P0, normal);
            double axisDistance = Math.Abs(aOffset - bOffset);

            double existingThicknessInternal = existingWall.ThicknessInternal > 0
                ? existingWall.ThicknessInternal
                : 0;

            double halfSumThickness = (candidateThicknessInternal + existingThicknessInternal) / 2.0;

            double penetrationDepth = halfSumThickness - axisDistance;
            if (penetrationDepth <= tol)
                return false;

            double aT0 = Dot2D(a0, aDir);
            double aT1 = Dot2D(a1, aDir);
            double aFrom = Math.Min(aT0, aT1);
            double aTo = Math.Max(aT0, aT1);

            double bT0 = Dot2D(existingWall.P0, aDir);
            double bT1 = Dot2D(existingWall.P1, aDir);
            double bFrom = Math.Min(bT0, bT1);
            double bTo = Math.Max(bT0, bT1);

            double overlapFrom = Math.Max(aFrom, bFrom);
            double overlapTo = Math.Min(aTo, bTo);

            return (overlapTo - overlapFrom) > tol;
        }

        private static XYZ Normalize2D(XYZ v)
        {
            if (v == null)
                return null;

            double len = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (len < 1e-12)
                return null;

            return new XYZ(v.X / len, v.Y / len, 0);
        }

        private static XYZ CanonicalizeDirection(XYZ dir)
        {
            if (dir == null)
                return null;

            if (dir.X < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            if (Math.Abs(dir.X) < 1e-9 && dir.Y < -1e-9)
                return new XYZ(-dir.X, -dir.Y, 0);

            return new XYZ(dir.X, dir.Y, 0);
        }

        private static double Dot2D(XYZ a, XYZ b)
        {
            return a.X * b.X + a.Y * b.Y;
        }

        private static double Cross2D(XYZ a, XYZ b)
        {
            return a.X * b.Y - a.Y * b.X;
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static bool PointOnSegment2D(XYZ p, XYZ a, XYZ b, double tol)
        {
            double ab = Distance2D(a, b);
            double ap = Distance2D(a, p);
            double pb = Distance2D(p, b);
            return Math.Abs((ap + pb) - ab) <= tol;
        }

        private static bool AreCollinear2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, double tol)
        {
            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            if (Math.Abs(Cross2D(ad, bd)) > tol)
                return false;

            if (Math.Abs(Cross2D(b0 - a0, ad)) > tol)
                return false;

            return true;
        }

        private static bool TryIntersectInfiniteLines2D(XYZ p1, XYZ d1, XYZ p2, XYZ d2, out XYZ intersection)
        {
            const double tol = 1e-12;
            intersection = null;

            double cross = Cross2D(d1, d2);
            if (Math.Abs(cross) < tol)
                return false;

            XYZ delta = p2 - p1;
            double t = Cross2D(delta, d2) / cross;

            intersection = new XYZ(
                p1.X + d1.X * t,
                p1.Y + d1.Y * t,
                0);

            return true;
        }

        private static bool TryIntersectSegments2D(XYZ a0, XYZ a1, XYZ b0, XYZ b1, out XYZ intersection, double tol)
        {
            intersection = null;

            XYZ ad = Normalize2D(a1 - a0);
            XYZ bd = Normalize2D(b1 - b0);

            if (ad == null || bd == null)
                return false;

            XYZ inter;
            if (!TryIntersectInfiniteLines2D(a0, ad, b0, bd, out inter))
                return false;

            if (!PointOnSegment2D(inter, a0, a1, tol))
                return false;

            if (!PointOnSegment2D(inter, b0, b1, tol))
                return false;

            intersection = inter;
            return true;
        }

        #endregion
    }
}