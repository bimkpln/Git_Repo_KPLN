using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Tools.Common;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

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
            preparedAxisLines = RemoveSegmentsOverlappingExistingWalls(preparedAxisLines, existingWalls);
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
                    debugMessages.Add("Не удалось удалить 2D-квартиру ID = " + GetElementIdValue(apartmentId) + ": " + ex.Message);

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
                bic = (BuiltInCategory)category.Id.IntegerValue;
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
                            debugMessages.Add("У вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + " не найден тип.");
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
                            debugMessages.Add("У вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + " не найден Transform.");
                        continue;
                    }

                    XYZ insertPoint = tr.Origin;
                    if (insertPoint == null)
                        continue;

                    Level level = ResolvePlacementLevelForNestedInstance(doc, nestedFi, apartmentFi);
                    if (level == null)
                    {
                        if (debugMessages != null)
                            debugMessages.Add("Не найден уровень для вложенного элемента ID = " + GetElementIdValue(nestedFi.Id));
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
                            debugMessages.Add("Не удалось создать экземпляр для вложенного элемента ID = " + GetElementIdValue(nestedFi.Id));
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
                }
                catch (Exception ex)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Ошибка копирования вложенного элемента ID = " + GetElementIdValue(nestedFi.Id) + ": " + ex.Message);
                }
            }
        }













        private static void DeleteIsolatedCreatedWalls(Document doc, List<Wall> createdWallsForApartment)
        {
            if (doc == null || createdWallsForApartment == null || createdWallsForApartment.Count == 0)
                return;

            double tol = ConvertMmToInternal(10);
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

            double maxDistanceToWallAxis = ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                preparedDoor.RequiresOppositeDoorTypeAfterWallFlip = false;

                if (preparedDoor == null || preparedDoor.InsertPoint == null || preparedDoor.RelatedRoom2D == null)
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




        private int PlaceDoorsForApartment(Document doc, PreparedApartmentDoors apartmentDoors, List<Wall> createdWallsForApartment, Level baseLevel,
            List<ElementId> createdDoorIds = null)
        {
            if (doc == null || apartmentDoors == null || createdWallsForApartment == null || baseLevel == null)
                return 0;

            if (apartmentDoors.Doors == null || apartmentDoors.Doors.Count == 0)
                return 0;

            int installedCount = 0;
            double maxDistanceToWallAxis = ConvertMmToInternal(500);

            foreach (PreparedDoorPlacement preparedDoor in apartmentDoors.Doors)
            {
                if (preparedDoor == null || preparedDoor.DoorSymbol == null || preparedDoor.InsertPoint == null)
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

                FamilySymbol symbolToPlace = preparedDoor.DoorSymbol;

                if (preparedDoor.RequiresOppositeDoorTypeAfterWallFlip)
                    symbolToPlace = GetOppositeDoorSymbol(doc, symbolToPlace);

                if (symbolToPlace == null)
                    continue;

                if (!symbolToPlace.IsActive)
                {
                    symbolToPlace.Activate();
                    doc.Regenerate();
                }

                FamilyInstance createdDoor = doc.Create.NewFamilyInstance(
                    projectedPoint,
                    symbolToPlace,
                    hostWall,
                    baseLevel,
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                if (createdDoor != null)
                {
                    installedCount++;

                    if (createdDoorIds != null)
                        createdDoorIds.Add(createdDoor.Id);
                }
            }

            return installedCount;
        }

        private bool TryFindBestHostWallForDoor(XYZ doorPoint, List<Wall> candidateWalls, double maxDistanceToWallAxis, out Wall bestWall, out XYZ bestProjectedPoint, out double bestDistance)
        {
            bestWall = null;
            bestProjectedPoint = null;
            bestDistance = double.MaxValue;

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

                if (distance > maxDistanceToWallAxis)
                    continue;

                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestWall = wall;
                    bestProjectedPoint = projectedPoint;
                }
            }

            return bestWall != null && bestProjectedPoint != null;
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
            double z = a.Z + (b.Z - a.Z) * t;

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

            IEnumerable<Wall> walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>();

            foreach (Wall wall in walls)
            {
                if (wall == null)
                    continue;

                if (wall.LevelId != levelId)
                    continue;

                LocationCurve lc = wall.Location as LocationCurve;
                if (lc == null)
                    continue;

                Line line = lc.Curve as Line;
                if (line == null || line.Length < 1e-9)
                    continue;

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
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

        private PreparedApartmentDoors PrepareDoorsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, List<string> debugMessages)
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

                int widthMm;
                if (!TryGetDoorWidthMmFrom2DTypeName(typeName, out widthMm) || widthMm <= 0)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не удалось определить ширину 2D-двери у экземпляра ID = " + GetElementIdValue(doorFi.Id));
                    continue;
                }

                string roomCategory = GetCommentsValue(doorFi);
                if (string.IsNullOrWhiteSpace(roomCategory))
                    roomCategory = "-";

                string presetKey = ApartmentDoorRequirementOption.BuildKey(roomCategory, typeName, widthMm);

                string selectedDoorTypeName = null;
                if (preset != null && preset.DoorsByRoomCategory != null)
                    preset.DoorsByRoomCategory.TryGetValue(presetKey, out selectedDoorTypeName);

                if (string.IsNullOrWhiteSpace(selectedDoorTypeName) ||
                    string.Equals(selectedDoorTypeName, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                {
                    if (debugMessages != null)
                        debugMessages.Add("Для двери [" + roomCategory + "] (" + widthMm + ") не выбран тип двери проекта.");
                    continue;
                }

                FamilySymbol baseDoorSymbol = FindDoorSymbolByDisplayNameAndWidth(doc, selectedDoorTypeName, widthMm);
                if (baseDoorSymbol == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не найден тип двери проекта '" + selectedDoorTypeName + "' с шириной " + widthMm + " мм.");
                    continue;
                }

                FamilySymbol resolvedDoorSymbol = ResolveDoorSymbolForPlacement(doc, doorFi, baseDoorSymbol, debugMessages);
                if (resolvedDoorSymbol == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add(
                            "Не удалось определить итоговый тип 3D-двери для 2D-двери ID = " +
                            GetElementIdValue(doorFi.Id) + ".");
                    continue;
                }

                Transform doorTransform = doorFi.GetTransform();
                if (doorTransform == null)
                {
                    if (debugMessages != null)
                        debugMessages.Add("Не удалось получить Transform для 2D-двери ID = " + GetElementIdValue(doorFi.Id));
                    continue;
                }

                XYZ insertPointInProject = doorTransform.Origin;

                FamilyInstance matchedRoom = FindBestMatchingRoomForDoor(
                    apartmentFi,
                    roomCategory,
                    insertPointInProject,
                    doc);

                XYZ expectedRoomPoint = matchedRoom != null
                    ? GetRoomCenterPoint(matchedRoom)
                    : null;

                result.Doors.Add(new PreparedDoorPlacement
                {
                    ApartmentId = apartmentFi.Id,
                    Door2DId = doorFi.Id,
                    RoomCategory = roomCategory,
                    DoorWidthMm = widthMm,
                    SelectedDoorTypeName = BuildDoorTypeDisplayName(resolvedDoorSymbol),
                    DoorSymbol = resolvedDoorSymbol,
                    InsertPoint = insertPointInProject,
                    RelatedRoom2D = matchedRoom,
                    RequiresOppositeDoorTypeAfterWallFlip = false
                });
            }

            return result;
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

                bool isGenericModel =
                    string.Equals(categoryName, "Обобщенные модели", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(categoryName, "Обобщённые модели", StringComparison.OrdinalIgnoreCase);

                bool isDoorFamily =
                    string.Equals(familyName, "Дверь", StringComparison.OrdinalIgnoreCase);

                if (isGenericModel && isDoorFamily)
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
                            GetElementIdValue(roomFi.Id) + ": " + ex.Message);
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
                        roomLevel.Elevation + ConvertMmToInternal(100));

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
                        double expectedAreaSquareMeters = ConvertInternalAreaToSquareMeters(preparedRoom.ExpectedAreaInternal);
                        double actualAreaSquareMeters = ConvertInternalAreaToSquareMeters(actualAreaInternal);

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

        private static List<Line> RemoveSegmentsOverlappingExistingWalls(List<Line> newLines, List<ExistingWallLineInfo> existingWalls)
        {
            const double tol = 1e-6;
            List<Line> result = new List<Line>();

            foreach (Line newLine in newLines)
            {
                if (newLine == null || newLine.Length <= tol)
                    continue;

                List<Line> remaining = SubtractExistingWallsFromNewLine(newLine, existingWalls);

                foreach (Line part in remaining)
                {
                    if (part != null && part.Length > tol)
                        result.Add(part);
                }
            }

            return result;
        }

        private static List<Line> SubtractExistingWallsFromNewLine(Line newLine, List<ExistingWallLineInfo> existingWalls)
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
                if (Math.Abs(exOffset - newOffset) > tol)
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
