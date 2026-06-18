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
        private int PlaceWindowGeometryInTransaction(Document doc, List<PreparedApartmentWalls> preparedApartments,
            List<PreparedApartmentWindows> preparedWindowsByApartment, Dictionary<long, List<Wall>> doorHostWallsByApartment,
            List<ExistingWallLineInfo> existingWalls, Level baseLevel, List<string> debugMessages,
            Dictionary<long, ApartmentProcessState> apartmentStates)
        {
            if (doc == null || preparedApartments == null || preparedApartments.Count == 0 || baseLevel == null)
                return 0;

            int installedWindowsCount = 0;

            using (Transaction t = new Transaction(doc, "KPLN. Менеджер квартир. Геометрия окон"))
            {
                t.Start();

                foreach (PreparedApartmentWalls apartmentWalls in preparedApartments)
                {
                    if (apartmentWalls == null)
                        continue;

                    PreparedApartmentWindows apartmentWindows = preparedWindowsByApartment != null
                        ? preparedWindowsByApartment.FirstOrDefault(x => x != null && x.ApartmentId == apartmentWalls.ApartmentId)
                        : null;

                    if (apartmentWindows == null || apartmentWindows.Windows == null || apartmentWindows.Windows.Count == 0)
                        continue;

                    ApartmentProcessState state = GetOrCreateApartmentState(apartmentStates, apartmentWalls.ApartmentId);
                    List<Wall> createdDoorHostWallsForApartment = GetCreatedWallsForApartment(doorHostWallsByApartment, apartmentWalls.ApartmentId);
                    List<ElementId> createdWindowIds = new List<ElementId>();

                    int installedWindowsForApartment = PlaceWindowsForApartment(
                        doc,
                        apartmentWindows,
                        createdDoorHostWallsForApartment,
                        existingWalls,
                        baseLevel,
                        debugMessages,
                        state,
                        createdWindowIds);

                    foreach (ElementId createdWindowId in createdWindowIds)
                        AddCreatedElementCandidate(state, createdWindowId);

                    installedWindowsCount += installedWindowsForApartment;

                    if (installedWindowsForApartment > 0)
                        state.HasInstalledWindows = true;

                    int skippedWindowsForApartment = apartmentWindows.Windows.Count - installedWindowsForApartment;
                    if (skippedWindowsForApartment > 0)
                        state.SkippedWindowsCount += skippedWindowsForApartment;
                }

                doc.Regenerate();
                t.Commit();
            }

            return installedWindowsCount;
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

                projectedPoint = AlignHostedFamilyInsertionPointToHostWallBase(projectedPoint, hostWall, baseLevel);

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

            bool useLevelHostOverload = ShouldCreateHostedFamilyWithLevelHostOverload(hostWall, baseLevel);
            if (!useLevelHostOverload && refDir != null)
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
            else if (useLevelHostOverload && debugMessages != null)
            {
                debugMessages.Add(
                    "Вставка окна: referenceDirection-overload пропущен, потому что host-стена не на нулевой отметке; используется host+level вставка. " +
                    "Точка = " + FormatPointMm(projectedPoint) +
                    ", уровень = " + FormatLevelDebugText(baseLevel) + ".");
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

        private PreparedApartmentWindows PrepareWindowsForApartment(Document doc, FamilyInstance apartmentFi, ApartmentPresetData preset, double placementPointZ,
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

                XYZ p0 = WithZ(marker.LocalP0, placementPointZ);
                XYZ p1 = WithZ(marker.LocalP1, placementPointZ);

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
    }
}