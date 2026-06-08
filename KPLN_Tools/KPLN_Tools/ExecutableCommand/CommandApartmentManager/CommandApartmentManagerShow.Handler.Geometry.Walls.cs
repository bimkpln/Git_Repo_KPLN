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
        private class CreatedApartmentWallGeometry
        {
            public Dictionary<long, List<Wall>> CreatedWallsByApartment { get; private set; }
            public Dictionary<long, List<Wall>> DoorHostWallsByApartment { get; private set; }

            public CreatedApartmentWallGeometry()
            {
                CreatedWallsByApartment = new Dictionary<long, List<Wall>>();
                DoorHostWallsByApartment = new Dictionary<long, List<Wall>>();
            }
        }

        private class ShaftWallAxisSplitResult
        {
            public List<Line> RegularAxisLines { get; set; }
            public List<Line> ShaftAxisLines { get; set; }
            public HashSet<string> MatchedMarkerKeys { get; set; }
        }

        private class ShaftMarkerLineData
        {
            public string Key { get; set; }
            public Line Line { get; set; }
        }

        private CreatedApartmentWallGeometry CreateWallGeometryInTransaction(Document doc, List<PreparedApartmentWalls> preparedApartments,
            Level baseLevel, Level topLevel, double baseOffsetInternal, double wallHeightInternal, List<ExistingWallLineInfo> existingWalls,
            double connectTol, Dictionary<long, ApartmentProcessState> apartmentStates)
        {
            CreatedApartmentWallGeometry result = new CreatedApartmentWallGeometry();
            if (doc == null || preparedApartments == null || preparedApartments.Count == 0 || baseLevel == null)
                return result;

            using (Transaction t = new Transaction(doc, "KPLN. Геометрия стен"))
            {
                t.Start();

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null || apartmentWalls.WallType == null || apartmentWalls.AxisLines == null)
                        continue;

                    long apartmentKey = GetApartmentGeometryKey(apartmentWalls.ApartmentId);
                    ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);
                    List<Wall> createdWallsForApartment = new List<Wall>();
                    List<Wall> createdDoorHostWallsForApartment = new List<Wall>();

                    foreach (Line axis in apartmentWalls.AxisLines)
                    {
                        if (axis == null || axis.Length < 1e-6)
                            continue;

                        Wall wall = Wall.Create(doc, axis, apartmentWalls.WallType.Id, baseLevel.Id, wallHeightInternal, 0, false, false);
                        ApplyWallPresetParameters(wall, baseLevel, topLevel, baseOffsetInternal, wallHeightInternal);
                        createdWallsForApartment.Add(wall);
                        createdDoorHostWallsForApartment.Add(wall);

                        AddCreatedElementCandidate(state, wall.Id);
                    }

                    if (apartmentWalls.ShaftWallType != null && apartmentWalls.ShaftAxisLines != null)
                    {
                        List<ExistingWallLineInfo> shaftSnapTargets = new List<ExistingWallLineInfo>();
                        if (existingWalls != null)
                            shaftSnapTargets.AddRange(existingWalls);

                        shaftSnapTargets.AddRange(BuildWallLineInfoFromWalls(createdWallsForApartment));

                        List<Line> shaftAxisLinesToCreate = SnapNewLinesToExistingWalls(apartmentWalls.ShaftAxisLines, shaftSnapTargets, connectTol);
                        shaftAxisLinesToCreate = MergeCollinearLines(shaftAxisLinesToCreate);

                        foreach (Line axis in shaftAxisLinesToCreate)
                        {
                            if (axis == null || axis.Length < 1e-6)
                                continue;

                            Wall wall = Wall.Create(doc, axis, apartmentWalls.ShaftWallType.Id, baseLevel.Id, wallHeightInternal, 0, false, false);
                            ApplyWallPresetParameters(wall, baseLevel, topLevel, baseOffsetInternal, wallHeightInternal);
                            createdWallsForApartment.Add(wall);

                            AddCreatedElementCandidate(state, wall.Id);
                        }
                    }

                    if (apartmentWalls.LoggiaWallType != null && apartmentWalls.LoggiaAxisLines != null)
                    {
                        List<ExistingWallLineInfo> loggiaSnapTargets = new List<ExistingWallLineInfo>();
                        if (existingWalls != null)
                            loggiaSnapTargets.AddRange(existingWalls);

                        loggiaSnapTargets.AddRange(BuildWallLineInfoFromWalls(createdWallsForApartment));

                        List<Line> loggiaAxisLinesToCreate = SnapNewLinesToExistingWalls(apartmentWalls.LoggiaAxisLines, loggiaSnapTargets, connectTol);
                        loggiaAxisLinesToCreate = MergeCollinearLines(loggiaAxisLinesToCreate);

                        foreach (Line axis in loggiaAxisLinesToCreate)
                        {
                            if (axis == null || axis.Length < 1e-6)
                                continue;

                            Wall wall = Wall.Create(doc, axis, apartmentWalls.LoggiaWallType.Id, baseLevel.Id, wallHeightInternal, 0, false, false);
                            ApplyWallPresetParameters(wall, baseLevel, topLevel, baseOffsetInternal, wallHeightInternal);
                            createdWallsForApartment.Add(wall);

                            AddCreatedElementCandidate(state, wall.Id);
                        }
                    }

                    if (createdWallsForApartment.Count > 0)
                        state.HasCreatedWalls = true;

                    result.CreatedWallsByApartment[apartmentKey] = createdWallsForApartment;
                    result.DoorHostWallsByApartment[apartmentKey] = createdDoorHostWallsForApartment;
                }

                doc.Regenerate();
                t.Commit();
            }

            return result;
        }

        private static long GetApartmentGeometryKey(ElementId apartmentId)
        {
            return apartmentId != null ? IDHelper.ElIdValue(apartmentId) : -1;
        }

        private static List<Wall> GetCreatedWallsForApartment(Dictionary<long, List<Wall>> wallsByApartment, ElementId apartmentId)
        {
            if (wallsByApartment == null)
                return new List<Wall>();

            List<Wall> walls;
            return wallsByApartment.TryGetValue(GetApartmentGeometryKey(apartmentId), out walls) && walls != null
                ? walls
                : new List<Wall>();
        }

        private List<Line> BuildPreparedWallAxisLinesForSingleRoom(FamilyInstance roomFi, double apartmentWallThicknessInternal, List<ExistingWallLineInfo> existingWalls,
            View geometryView, double connectTol, double intersectionTol, List<string> debugMessages, ref int skippedWallsForApartment)
        {
            if (roomFi == null)
                return new List<Line>();

            CurveLoop roomLoop = BuildRoomLoopFromInstance(roomFi, geometryView, debugMessages);
            if (roomLoop == null)
                return new List<Line>();

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

        private static Line GetWallAxisLine(Wall wall)
        {
            if (wall == null)
                return null;

            LocationCurve lc = wall.Location as LocationCurve;
            if (lc == null)
                return null;

            return lc.Curve as Line;
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

        private static ShaftWallAxisSplitResult SplitApartmentWallAxesByShaftMarkers(List<Line> apartmentAxisLines,
            List<FamilyShaftWallMarker> shaftMarkers, double apartmentWallThicknessInternal, WallType shaftWallType)
        {
            ShaftWallAxisSplitResult result = new ShaftWallAxisSplitResult
            {
                RegularAxisLines = new List<Line>(),
                ShaftAxisLines = new List<Line>(),
                MatchedMarkerKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            };

            if (apartmentAxisLines == null || apartmentAxisLines.Count == 0)
                return result;

            List<ShaftMarkerLineData> markerLines = new List<ShaftMarkerLineData>();
            if (shaftMarkers != null)
            {
                foreach (FamilyShaftWallMarker marker in shaftMarkers)
                {
                    if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                        continue;

                    if (Distance2D(marker.ProjectP0, marker.ProjectP1) < IDHelper.ConvertMmToInternal(10))
                        continue;

                    markerLines.Add(new ShaftMarkerLineData
                    {
                        Key = BuildLineMarkerKey(marker.ProjectP0, marker.ProjectP1),
                        Line = Line.CreateBound(marker.ProjectP0, marker.ProjectP1)
                    });
                }
            }

            if (markerLines.Count == 0 || shaftWallType == null)
            {
                result.RegularAxisLines = MergeCollinearLines(apartmentAxisLines);
                return result;
            }

            const double tol = 1e-6;
            double shaftWallThicknessInternal = GetWallTypeWidthInternal(shaftWallType);
            double markerOffsetTol =
                Math.Max(apartmentWallThicknessInternal, shaftWallThicknessInternal) / 2.0 +
                IDHelper.ConvertMmToInternal(80);
            double minOverlap = IDHelper.ConvertMmToInternal(20);

            foreach (Line axis in apartmentAxisLines)
            {
                if (axis == null || axis.Length <= tol)
                    continue;

                XYZ p0 = axis.GetEndPoint(0);
                XYZ p1 = axis.GetEndPoint(1);
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

                List<Interval1D> shaftIntervals = new List<Interval1D>();

                foreach (ShaftMarkerLineData markerLine in markerLines)
                {
                    Interval1D overlap;
                    if (markerLine == null ||
                        markerLine.Line == null ||
                        !TryGetShaftMarkerOverlapInterval(axis, markerLine.Line, markerOffsetTol, minOverlap, out overlap))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(markerLine.Key))
                        result.MatchedMarkerKeys.Add(markerLine.Key);

                    if (overlap != null)
                        shaftIntervals.Add(overlap);
                }

                if (shaftIntervals.Count == 0)
                {
                    result.RegularAxisLines.Add(axis);
                    continue;
                }

                List<Interval1D> mergedShaftIntervals = MergeIntervals(shaftIntervals);
                foreach (Interval1D shaftInterval in mergedShaftIntervals)
                {
                    double shaftFrom = Math.Max(from, shaftInterval.From);
                    double shaftTo = Math.Min(to, shaftInterval.To);
                    if (shaftTo - shaftFrom <= tol)
                        continue;

                    Line shaftAxis = BuildGenericAxisLine(dir, offset, z, shaftFrom, shaftTo);
                    if (shaftAxis != null && shaftAxis.Length > tol)
                        result.ShaftAxisLines.Add(shaftAxis);
                }

                List<Interval1D> regularIntervals = SubtractIntervals(
                    new Interval1D { From = from, To = to },
                    mergedShaftIntervals);

                foreach (Interval1D regularInterval in regularIntervals)
                {
                    if (regularInterval == null || regularInterval.To - regularInterval.From <= tol)
                        continue;

                    Line regularAxis = BuildGenericAxisLine(dir, offset, z, regularInterval.From, regularInterval.To);
                    if (regularAxis != null && regularAxis.Length > tol)
                        result.RegularAxisLines.Add(regularAxis);
                }
            }

            result.RegularAxisLines = MergeCollinearLines(result.RegularAxisLines);
            result.ShaftAxisLines = MergeCollinearLines(result.ShaftAxisLines);
            return result;
        }

        private static bool TryGetShaftMarkerOverlapInterval(Line axisLine, Line markerLine, double offsetTolerance, double minOverlap,
            out Interval1D overlap)
        {
            overlap = null;

            if (axisLine == null || markerLine == null)
                return false;

            XYZ a0 = axisLine.GetEndPoint(0);
            XYZ a1 = axisLine.GetEndPoint(1);
            XYZ m0 = markerLine.GetEndPoint(0);
            XYZ m1 = markerLine.GetEndPoint(1);

            XYZ axisDir = Normalize2D(a1 - a0);
            XYZ markerDir = Normalize2D(m1 - m0);
            if (axisDir == null || markerDir == null)
                return false;

            axisDir = CanonicalizeDirection(axisDir);
            markerDir = CanonicalizeDirection(markerDir);

            if (Math.Abs(Dot2D(axisDir, markerDir)) < 0.98)
                return false;

            XYZ normal = new XYZ(-axisDir.Y, axisDir.X, 0);
            double axisOffset = Dot2D(a0, normal);
            double markerOffset0 = Dot2D(m0, normal);
            double markerOffset1 = Dot2D(m1, normal);
            if (Math.Max(Math.Abs(axisOffset - markerOffset0), Math.Abs(axisOffset - markerOffset1)) > offsetTolerance)
                return false;

            double axisT0 = Dot2D(a0, axisDir);
            double axisT1 = Dot2D(a1, axisDir);
            double axisFrom = Math.Min(axisT0, axisT1);
            double axisTo = Math.Max(axisT0, axisT1);

            double markerT0 = Dot2D(m0, axisDir);
            double markerT1 = Dot2D(m1, axisDir);
            double markerFrom = Math.Min(markerT0, markerT1);
            double markerTo = Math.Max(markerT0, markerT1);

            double overlapFrom = Math.Max(axisFrom, markerFrom);
            double overlapTo = Math.Min(axisTo, markerTo);

            if (overlapTo - overlapFrom < minOverlap)
                return false;

            overlap = new Interval1D
            {
                From = overlapFrom,
                To = overlapTo
            };

            return true;
        }

        private static List<Line> BuildShaftWallAxesFromUnmatchedMarkers(List<FamilyShaftWallMarker> shaftMarkers,
            HashSet<string> matchedMarkerKeys, WallType shaftWallType, List<FamilyShaftFillRegion> shaftFillRegions,
            out int skippedMarkers)
        {
            skippedMarkers = 0;
            List<Line> result = new List<Line>();
            if (shaftMarkers == null || shaftMarkers.Count == 0 || shaftWallType == null)
                return result;

            foreach (FamilyShaftWallMarker marker in shaftMarkers)
            {
                if (marker == null || marker.ProjectP0 == null || marker.ProjectP1 == null)
                    continue;

                if (Distance2D(marker.ProjectP0, marker.ProjectP1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                string key = BuildLineMarkerKey(marker.ProjectP0, marker.ProjectP1);
                if (matchedMarkerKeys != null && matchedMarkerKeys.Contains(key))
                    continue;

                Line faceLine = Line.CreateBound(marker.ProjectP0, marker.ProjectP1);
                Line axisLine = BuildShaftWallAxisIntoFillRegion(faceLine, shaftWallType, shaftFillRegions);
                if (axisLine != null && axisLine.Length > 1e-6)
                    result.Add(axisLine);
                else
                    skippedMarkers++;
            }

            return MergeCollinearLines(result);
        }

        private static Line BuildShaftWallAxisIntoFillRegion(Line faceLine, WallType wallType, List<FamilyShaftFillRegion> shaftFillRegions)
        {
            if (faceLine == null || wallType == null)
                return null;

            XYZ p0 = faceLine.GetEndPoint(0);
            XYZ p1 = faceLine.GetEndPoint(1);
            XYZ dir = Normalize2D(p1 - p0);
            if (dir == null)
                return null;

            double wallWidth = GetWallTypeWidthInternal(wallType);
            if (wallWidth <= 1e-9)
                return null;

            XYZ normal = new XYZ(-dir.Y, dir.X, 0);
            FamilyShaftFillRegion fillRegion = FindShaftFillRegionForFaceLine(faceLine, shaftFillRegions);
            if (fillRegion == null || fillRegion.Boundary == null || fillRegion.Boundary.Count < 3)
                return null;

            XYZ interiorPoint = TryGetPolygonCentroid(fillRegion.Boundary);
            if (interiorPoint == null)
                interiorPoint = GetAveragePoint(fillRegion.Boundary);

            if (interiorPoint == null)
                return null;

            XYZ mid = new XYZ(
                0.5 * (p0.X + p1.X),
                0.5 * (p0.Y + p1.Y),
                0.5 * (p0.Z + p1.Z));

            XYZ toInterior = Normalize2D(interiorPoint - mid);
            if (toInterior == null)
                return null;

            XYZ interiorDirection = Dot2D(toInterior, normal) >= 0
                ? normal
                : new XYZ(-normal.X, -normal.Y, 0);

            double halfWidth = wallWidth / 2.0;
            XYZ offset = new XYZ(interiorDirection.X * halfWidth, interiorDirection.Y * halfWidth, 0);

            return Line.CreateBound(
                new XYZ(p0.X + offset.X, p0.Y + offset.Y, p0.Z),
                new XYZ(p1.X + offset.X, p1.Y + offset.Y, p1.Z));
        }

        private static FamilyShaftFillRegion FindShaftFillRegionForFaceLine(Line faceLine, List<FamilyShaftFillRegion> shaftFillRegions)
        {
            if (faceLine == null || shaftFillRegions == null || shaftFillRegions.Count == 0)
                return null;

            FamilyShaftFillRegion bestRegion = null;
            double bestArea = double.PositiveInfinity;
            foreach (FamilyShaftFillRegion region in shaftFillRegions)
            {
                if (region == null || region.Boundary == null || region.Boundary.Count < 3)
                    continue;

                if (!ShaftFaceLineMatchesFillRegionEdge(faceLine, region.Boundary))
                    continue;

                double area = Math.Abs(GetSignedAreaXY(region.Boundary));
                if (area <= IDHelper.ConvertMmToInternal(100) * IDHelper.ConvertMmToInternal(100))
                    continue;

                if (area < bestArea)
                {
                    bestArea = area;
                    bestRegion = region;
                }
            }

            return bestRegion;
        }

        private static bool ShaftFaceLineMatchesFillRegionEdge(Line faceLine, List<XYZ> boundary)
        {
            if (faceLine == null || boundary == null || boundary.Count < 3)
                return false;

            XYZ f0 = faceLine.GetEndPoint(0);
            XYZ f1 = faceLine.GetEndPoint(1);
            XYZ faceDir = Normalize2D(f1 - f0);
            if (faceDir == null)
                return false;

            double faceLength = Distance2D(f0, f1);
            double collinearTol = IDHelper.ConvertMmToInternal(25);
            double minOverlap = faceLength * 0.7;

            for (int i = 0; i < boundary.Count; i++)
            {
                XYZ e0 = boundary[i];
                XYZ e1 = boundary[(i + 1) % boundary.Count];
                if (e0 == null || e1 == null || Distance2D(e0, e1) < IDHelper.ConvertMmToInternal(10))
                    continue;

                if (!AreCollinear2D(f0, f1, e0, e1, collinearTol))
                    continue;

                double fT0 = Dot2D(f0, faceDir);
                double fT1 = Dot2D(f1, faceDir);
                double fFrom = Math.Min(fT0, fT1);
                double fTo = Math.Max(fT0, fT1);

                double eT0 = Dot2D(e0, faceDir);
                double eT1 = Dot2D(e1, faceDir);
                double eFrom = Math.Min(eT0, eT1);
                double eTo = Math.Max(eT0, eT1);

                double overlap = Math.Min(fTo, eTo) - Math.Max(fFrom, eFrom);
                if (overlap >= minOverlap)
                    return true;
            }

            return false;
        }

        private static XYZ GetAveragePoint(List<XYZ> points)
        {
            if (points == null || points.Count == 0)
                return null;

            double x = 0;
            double y = 0;
            double z = 0;
            int count = 0;

            foreach (XYZ point in points)
            {
                if (point == null)
                    continue;

                x += point.X;
                y += point.Y;
                z += point.Z;
                count++;
            }

            return count > 0
                ? new XYZ(x / count, y / count, z / count)
                : null;
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
    }
}