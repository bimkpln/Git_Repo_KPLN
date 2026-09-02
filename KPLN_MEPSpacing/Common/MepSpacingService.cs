using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_MEPSpacing.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPSpacing.Common
{
    internal static class MepSpacingService
    {
        private const double Epsilon = 1e-9;
        private const double DirectionTolerance = 0.998;
        private const double VerticalDirectionTolerance = 0.985;
        private const double VerticalAlignmentToleranceMm = 25;
        private const double SolverRegularization = 1e-8;

        public static SpacingApplyResult ApplySpacing(Document doc, IEnumerable<ElementId> elementIds, IEnumerable<ElementId> baseElementIds, double distanceMm, SpacingCalculationMode mode)
        {
            if (doc == null)
                throw new InvalidOperationException("Документ Revit не найден.");

            double distance = MillimetersToInternal(distanceMm);
            List<ElementId> uniqueElementIds = (elementIds ?? Enumerable.Empty<ElementId>())
                .Where(id => id != null && !id.Equals(ElementId.InvalidElementId))
                .GroupBy(GetElementIdValue)
                .Select(group => group.First())
                .ToList();

            List<MepElementInfo> elements = uniqueElementIds
                .Select(id => doc.GetElement(id))
                .Where(element => element != null)
                .Select(CreateInfo)
                .ToList();

            if (elements.Count < 2)
                throw new InvalidOperationException("Выбери минимум два MEP-элемента.");

            List<MepRouteInfo> routes = BuildRoutes(doc, elements);
            if (routes.Count < 2)
                throw new InvalidOperationException("Не получилось разделить выбранные элементы минимум на две трассы.");

            HashSet<long> baseElementIdValues = new HashSet<long>((baseElementIds ?? Enumerable.Empty<ElementId>()).Select(GetElementIdValue));
            foreach (MepRouteInfo route in routes)
            {
                route.IsFixed = IsFixedRoute(doc, route, baseElementIdValues);
                if (!route.IsFixed)
                    continue;

                foreach (MepElementInfo element in route.Elements)
                    element.IsFixed = true;
            }

            SpacingApplyResult result = new SpacingApplyResult();
            Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement = elements.ToDictionary(element => element, _ => new List<MoveConstraint>());
            List<DirectionFamilyInfo> directionFamilies = BuildDirectionFamilies(elements);
            int processedRowCount = 0;

            foreach (DirectionFamilyInfo directionFamily in directionFamilies)
                processedRowCount += AddSpacingConstraintsForDirectionFamily(doc, routes, directionFamily, distance, mode, constraintsByElement);

            processedRowCount += AddSpacingConstraintsForVerticalElements(doc, routes, distance, mode, constraintsByElement);

            result.FixedElementCount = elements.Count(element => element.IsFixed);

            if (processedRowCount == 0)
                result.Messages.Add("Не найдено соседних параллельных трасс для расчёта шага.");

            MoveElements(doc, elements, constraintsByElement, result);
            return result;
        }

        private static int AddSpacingConstraintsForDirectionFamily(
            Document doc,
            IReadOnlyList<MepRouteInfo> routes,
            DirectionFamilyInfo directionFamily,
            double distance,
            SpacingCalculationMode mode,
            Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement)
        {
            List<SegmentOffset> segmentOffsets = routes
                .SelectMany(route => route.Elements
                    .Where(element => Math.Abs(element.Direction.DotProduct(directionFamily.Direction)) >= DirectionTolerance)
                    .Select(element => CreateSegmentOffset(doc, route, element, directionFamily.Normal)))
                .OrderBy(offset => offset.CurrentOffset)
                .ToList();

            if (segmentOffsets.Count < 2)
                return 0;

            int processedRowCount = 0;
            foreach (List<SegmentOffset> segmentRow in SplitIntoNeighbourRows(segmentOffsets, distance, mode))
            {
                if (AddSpacingConstraintsForSegmentRow(segmentRow, directionFamily.Normal, distance, mode, constraintsByElement))
                    processedRowCount++;
            }

            return processedRowCount;
        }

        private static int AddSpacingConstraintsForVerticalElements(
            Document doc,
            IReadOnlyList<MepRouteInfo> routes,
            double distance,
            SpacingCalculationMode mode,
            Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement)
        {
            int processedRowCount = 0;
            processedRowCount += AddSpacingConstraintsForVerticalAxis(doc, routes, XYZ.BasisX, XYZ.BasisY, distance, mode, constraintsByElement);
            processedRowCount += AddSpacingConstraintsForVerticalAxis(doc, routes, XYZ.BasisY, XYZ.BasisX, distance, mode, constraintsByElement);
            return processedRowCount;
        }

        private static int AddSpacingConstraintsForVerticalAxis(
            Document doc,
            IReadOnlyList<MepRouteInfo> routes,
            XYZ spacingDirection,
            XYZ alignmentDirection,
            double distance,
            SpacingCalculationMode mode,
            Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement)
        {
            List<SegmentOffset> offsets = routes
                .SelectMany(route => route.Elements
                    .Where(IsVerticalElement)
                    .Select(element => CreateSegmentOffset(doc, route, element, spacingDirection, alignmentDirection)))
                .OrderBy(offset => offset.AlignmentOffset)
                .ThenBy(offset => offset.CurrentOffset)
                .ToList();

            if (offsets.Count < 2)
                return 0;

            int processedRowCount = 0;
            double alignmentTolerance = MillimetersToInternal(VerticalAlignmentToleranceMm);
            foreach (List<SegmentOffset> alignedRow in SplitIntoAlignedRows(offsets, alignmentTolerance))
            {
                List<SegmentOffset> orderedAlignedRow = alignedRow
                    .OrderBy(offset => offset.CurrentOffset)
                    .ToList();

                foreach (List<SegmentOffset> segmentRow in SplitIntoNeighbourRows(orderedAlignedRow, distance, mode))
                {
                    if (AddSpacingConstraintsForSegmentRow(segmentRow, spacingDirection, distance, mode, constraintsByElement))
                        processedRowCount++;
                }
            }

            return processedRowCount;
        }

        private static bool AddSpacingConstraintsForSegmentRow(
            IReadOnlyList<SegmentOffset> segmentRow,
            XYZ spacingDirection,
            double distance,
            SpacingCalculationMode mode,
            Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement)
        {
            List<SegmentOffset> rowOffsets = segmentRow
                .OrderBy(offset => offset.CurrentOffset)
                .ToList();

            if (rowOffsets.Count < 2)
                return false;

            if (!rowOffsets.Any(offset => offset.ElementInfo.IsFixed))
                rowOffsets[0].ElementInfo.IsFixed = true;

            Dictionary<SegmentOffset, double> targetOffsets = BuildTargetOffsets(rowOffsets, distance, mode);
            foreach (SegmentOffset offset in rowOffsets.Where(offset => !offset.ElementInfo.IsFixed))
            {
                double delta = targetOffsets[offset] - offset.CurrentOffset;
                if (Math.Abs(delta) < Epsilon)
                    continue;

                constraintsByElement[offset.ElementInfo].Add(new MoveConstraint
                {
                    Direction = spacingDirection,
                    Delta = delta
                });
            }

            return true;
        }

        private static IEnumerable<List<SegmentOffset>> SplitIntoAlignedRows(IReadOnlyList<SegmentOffset> orderedOffsets, double alignmentTolerance)
        {
            List<SegmentOffset> currentRow = new List<SegmentOffset>();
            foreach (SegmentOffset offset in orderedOffsets)
            {
                if (currentRow.Count == 0)
                {
                    currentRow.Add(offset);
                    continue;
                }

                double averageAlignmentOffset = currentRow.Average(rowOffset => rowOffset.AlignmentOffset);
                if (Math.Abs(offset.AlignmentOffset - averageAlignmentOffset) > alignmentTolerance)
                {
                    yield return currentRow;
                    currentRow = new List<SegmentOffset>();
                }

                currentRow.Add(offset);
            }

            if (currentRow.Count > 0)
                yield return currentRow;
        }

        private static IEnumerable<List<SegmentOffset>> SplitIntoNeighbourRows(IReadOnlyList<SegmentOffset> orderedOffsets, double distance, SpacingCalculationMode mode)
        {
            List<SegmentOffset> currentRow = new List<SegmentOffset>();
            foreach (SegmentOffset offset in orderedOffsets)
            {
                if (currentRow.Count == 0)
                {
                    currentRow.Add(offset);
                    continue;
                }

                SegmentOffset previousOffset = currentRow[currentRow.Count - 1];
                double actualGap = Math.Abs(offset.CurrentOffset - previousOffset.CurrentOffset);
                double expectedGap = GetStepBetween(previousOffset, offset, distance, mode);
                double splitGap = Math.Max(expectedGap * 3, expectedGap + MillimetersToInternal(500));

                if (actualGap > splitGap)
                {
                    yield return currentRow;
                    currentRow = new List<SegmentOffset>();
                }

                currentRow.Add(offset);
            }

            if (currentRow.Count > 0)
                yield return currentRow;
        }

        private static List<RouteRowOffset> CreateRouteRowOffsets(IReadOnlyList<SegmentOffset> segmentRow)
        {
            return segmentRow
                .GroupBy(offset => offset.Route)
                .Select(group =>
                {
                    List<SegmentOffset> offsets = group.ToList();
                    double summaryLength = offsets.Sum(offset => offset.ElementInfo.Length);
                    double currentOffset = summaryLength > Epsilon
                        ? offsets.Sum(offset => offset.CurrentOffset * offset.ElementInfo.Length) / summaryLength
                        : offsets.Average(offset => offset.CurrentOffset);

                    return new RouteRowOffset
                    {
                        Route = group.Key,
                        CurrentOffset = currentOffset,
                        HalfSize = offsets.Select(offset => offset.HalfSize).DefaultIfEmpty(0).Max()
                    };
                })
                .ToList();
        }

        private static Dictionary<SegmentOffset, double> BuildTargetOffsets(IReadOnlyList<SegmentOffset> orderedOffsets, double distance, SpacingCalculationMode mode)
        {
            Dictionary<SegmentOffset, double> targetOffsets = new Dictionary<SegmentOffset, double>();
            List<int> fixedIndexes = orderedOffsets
                .Select((offset, index) => new { offset, index })
                .Where(item => item.offset.ElementInfo.IsFixed)
                .Select(item => item.index)
                .ToList();

            foreach (int fixedIndex in fixedIndexes)
                targetOffsets[orderedOffsets[fixedIndex]] = orderedOffsets[fixedIndex].CurrentOffset;

            int firstFixedIndex = fixedIndexes.First();
            for (int i = firstFixedIndex - 1; i >= 0; i--)
                targetOffsets[orderedOffsets[i]] = targetOffsets[orderedOffsets[i + 1]] - GetStepBetween(orderedOffsets[i], orderedOffsets[i + 1], distance, mode);

            for (int fixedIndexNumber = 0; fixedIndexNumber < fixedIndexes.Count; fixedIndexNumber++)
            {
                int fixedIndex = fixedIndexes[fixedIndexNumber];
                int nextFixedIndex = fixedIndexNumber + 1 < fixedIndexes.Count
                    ? fixedIndexes[fixedIndexNumber + 1]
                    : orderedOffsets.Count;

                for (int i = fixedIndex + 1; i < nextFixedIndex; i++)
                    targetOffsets[orderedOffsets[i]] = targetOffsets[orderedOffsets[i - 1]] + GetStepBetween(orderedOffsets[i - 1], orderedOffsets[i], distance, mode);
            }

            return targetOffsets;
        }

        private static double GetStepBetween(SegmentOffset firstOffset, SegmentOffset secondOffset, double distance, SpacingCalculationMode mode)
        {
            if (mode == SpacingCalculationMode.Centerline)
                return distance;

            return firstOffset.HalfSize + distance + secondOffset.HalfSize;
        }

        private static double GetStepBetween(RouteRowOffset firstOffset, RouteRowOffset secondOffset, double distance, SpacingCalculationMode mode)
        {
            if (mode == SpacingCalculationMode.Centerline)
                return distance;

            return firstOffset.HalfSize + distance + secondOffset.HalfSize;
        }

        private static void MoveElements(Document doc, IReadOnlyList<MepElementInfo> elements, Dictionary<MepElementInfo, List<MoveConstraint>> constraintsByElement, SpacingApplyResult result)
        {
            double tolerance = Math.Max(doc.Application.ShortCurveTolerance, MillimetersToInternal(0.5));
            foreach (MepElementInfo element in elements.Where(element => !element.IsFixed))
            {
                if (!constraintsByElement.TryGetValue(element, out List<MoveConstraint> constraints) || constraints.Count == 0)
                    continue;

                XYZ moveVector = GetMoveVector(constraints);
                if (moveVector.GetLength() < tolerance)
                    continue;

                TryMoveElement(doc, element, moveVector, result);
            }
        }

        private static XYZ GetMoveVector(IReadOnlyList<MoveConstraint> constraints)
        {
            double a11 = 0;
            double a12 = 0;
            double a22 = 0;
            double b1 = 0;
            double b2 = 0;

            foreach (MoveConstraint constraint in constraints)
            {
                XYZ direction = constraint.Direction;
                a11 += direction.X * direction.X;
                a12 += direction.X * direction.Y;
                a22 += direction.Y * direction.Y;
                b1 += direction.X * constraint.Delta;
                b2 += direction.Y * constraint.Delta;
            }

            double determinant = a11 * a22 - a12 * a12;
            if (Math.Abs(determinant) > Epsilon)
            {
                double x = (b1 * a22 - b2 * a12) / determinant;
                double y = (a11 * b2 - a12 * b1) / determinant;
                return new XYZ(x, y, 0);
            }

            XYZ weightedVector = XYZ.Zero;
            double weight = 0;
            foreach (MoveConstraint constraint in constraints)
            {
                weightedVector += constraint.Direction.Multiply(constraint.Delta);
                weight += constraint.Direction.DotProduct(constraint.Direction);
            }

            return weight > Epsilon ? weightedVector.Divide(weight) : XYZ.Zero;
        }

        private static void TryMoveElement(Document doc, MepElementInfo element, XYZ moveVector, SpacingApplyResult result)
        {
            SubTransaction subTransaction = new SubTransaction(doc);
            try
            {
                subTransaction.Start();
                ElementTransformUtils.MoveElement(doc, element.Element.Id, moveVector);
                subTransaction.Commit();
                result.MovedElementCount++;
            }
            catch (Exception ex)
            {
                try
                {
                    if (subTransaction.GetStatus() == TransactionStatus.Started)
                        subTransaction.RollBack();
                }
                catch
                {
                }

                result.SkippedElementCount++;
                result.Messages.Add($"Пропущен элемент id {GetElementIdValue(element.Element.Id)}: {ex.Message}");
            }
        }

        private static Dictionary<MepRouteInfo, XYZ> BuildRouteMoveVectors(Document doc, IReadOnlyList<MepRouteInfo> routes, IReadOnlyList<PairSpacingConstraint> constraints, SpacingApplyResult result)
        {
            List<MepRouteInfo> movingRoutes = routes.Where(route => !route.IsFixed).ToList();
            Dictionary<MepRouteInfo, XYZ> moveVectorsByRoute = routes.ToDictionary(route => route, _ => XYZ.Zero);
            if (movingRoutes.Count == 0 || constraints.Count == 0)
            {
                AddConstraintResidualWarnings(doc, constraints, moveVectorsByRoute, result);
                return moveVectorsByRoute;
            }

            Dictionary<MepRouteInfo, int> routeIndexes = movingRoutes
                .Select((route, index) => new { route, index })
                .ToDictionary(item => item.route, item => item.index);

            int variableCount = movingRoutes.Count * 2;
            double[,] matrix = new double[variableCount, variableCount];
            double[] rightSide = new double[variableCount];

            foreach (PairSpacingConstraint constraint in constraints)
            {
                List<EquationCoefficient> coefficients = new List<EquationCoefficient>();
                AddRouteCoefficients(constraint.FirstRoute, routeIndexes, -constraint.Direction.X, -constraint.Direction.Y, coefficients);
                AddRouteCoefficients(constraint.SecondRoute, routeIndexes, constraint.Direction.X, constraint.Direction.Y, coefficients);

                if (coefficients.Count == 0)
                    continue;

                foreach (EquationCoefficient firstCoefficient in coefficients)
                {
                    rightSide[firstCoefficient.Index] += firstCoefficient.Value * constraint.Delta;
                    foreach (EquationCoefficient secondCoefficient in coefficients)
                        matrix[firstCoefficient.Index, secondCoefficient.Index] += firstCoefficient.Value * secondCoefficient.Value;
                }
            }

            for (int i = 0; i < variableCount; i++)
            {
                matrix[i, i] += SolverRegularization;

                if (IsEmptyEquation(matrix, i))
                    matrix[i, i] = 1;
            }

            if (!TrySolveLinearSystem(matrix, rightSide, out double[] solution))
            {
                result.Messages.Add("Не удалось рассчитать общий вектор перемещения для трасс.");
                return moveVectorsByRoute;
            }

            foreach (MepRouteInfo route in movingRoutes)
            {
                int routeIndex = routeIndexes[route];
                moveVectorsByRoute[route] = new XYZ(solution[routeIndex * 2], solution[routeIndex * 2 + 1], 0);
            }

            AddConstraintResidualWarnings(doc, constraints, moveVectorsByRoute, result);
            return moveVectorsByRoute;
        }

        private static void AddRouteCoefficients(MepRouteInfo route, Dictionary<MepRouteInfo, int> routeIndexes, double xCoefficient, double yCoefficient, List<EquationCoefficient> coefficients)
        {
            if (!routeIndexes.TryGetValue(route, out int routeIndex))
                return;

            coefficients.Add(new EquationCoefficient
            {
                Index = routeIndex * 2,
                Value = xCoefficient
            });
            coefficients.Add(new EquationCoefficient
            {
                Index = routeIndex * 2 + 1,
                Value = yCoefficient
            });
        }

        private static bool IsEmptyEquation(double[,] matrix, int row)
        {
            for (int column = 0; column < matrix.GetLength(1); column++)
            {
                if (Math.Abs(matrix[row, column]) > Epsilon)
                    return false;
            }

            return true;
        }

        private static bool TrySolveLinearSystem(double[,] sourceMatrix, double[] sourceRightSide, out double[] solution)
        {
            int size = sourceRightSide.Length;
            double[,] matrix = (double[,])sourceMatrix.Clone();
            double[] rightSide = (double[])sourceRightSide.Clone();
            solution = new double[size];

            for (int pivot = 0; pivot < size; pivot++)
            {
                int bestRow = pivot;
                double bestValue = Math.Abs(matrix[pivot, pivot]);
                for (int row = pivot + 1; row < size; row++)
                {
                    double value = Math.Abs(matrix[row, pivot]);
                    if (value <= bestValue)
                        continue;

                    bestValue = value;
                    bestRow = row;
                }

                if (bestValue < Epsilon)
                    return false;

                if (bestRow != pivot)
                    SwapRows(matrix, rightSide, pivot, bestRow);

                double pivotValue = matrix[pivot, pivot];
                for (int column = pivot; column < size; column++)
                    matrix[pivot, column] /= pivotValue;
                rightSide[pivot] /= pivotValue;

                for (int row = 0; row < size; row++)
                {
                    if (row == pivot)
                        continue;

                    double factor = matrix[row, pivot];
                    if (Math.Abs(factor) < Epsilon)
                        continue;

                    for (int column = pivot; column < size; column++)
                        matrix[row, column] -= factor * matrix[pivot, column];
                    rightSide[row] -= factor * rightSide[pivot];
                }
            }

            for (int i = 0; i < size; i++)
                solution[i] = rightSide[i];

            return true;
        }

        private static void SwapRows(double[,] matrix, double[] rightSide, int firstRow, int secondRow)
        {
            int columnCount = matrix.GetLength(1);
            for (int column = 0; column < columnCount; column++)
            {
                double value = matrix[firstRow, column];
                matrix[firstRow, column] = matrix[secondRow, column];
                matrix[secondRow, column] = value;
            }

            double rightSideValue = rightSide[firstRow];
            rightSide[firstRow] = rightSide[secondRow];
            rightSide[secondRow] = rightSideValue;
        }

        private static void AddConstraintResidualWarnings(Document doc, IReadOnlyList<PairSpacingConstraint> constraints, Dictionary<MepRouteInfo, XYZ> moveVectorsByRoute, SpacingApplyResult result)
        {
            double tolerance = Math.Max(doc.Application.ShortCurveTolerance, MillimetersToInternal(1));
            HashSet<long> warnedRouteIds = new HashSet<long>();

            foreach (PairSpacingConstraint constraint in constraints)
            {
                XYZ firstMove = moveVectorsByRoute.TryGetValue(constraint.FirstRoute, out XYZ firstRouteMove) ? firstRouteMove : XYZ.Zero;
                XYZ secondMove = moveVectorsByRoute.TryGetValue(constraint.SecondRoute, out XYZ secondRouteMove) ? secondRouteMove : XYZ.Zero;
                double actualDelta = (secondMove - firstMove).DotProduct(constraint.Direction);
                if (Math.Abs(actualDelta - constraint.Delta) <= tolerance)
                    continue;

                AddRouteResidualWarning(constraint.FirstRoute, warnedRouteIds, result);
                AddRouteResidualWarning(constraint.SecondRoute, warnedRouteIds, result);
            }
        }

        private static void AddRouteResidualWarning(MepRouteInfo route, HashSet<long> warnedRouteIds, SpacingApplyResult result)
        {
            long routeId = GetElementIdValue(route.Elements.First().Element.Id);
            if (!warnedRouteIds.Add(routeId))
                return;

            result.Messages.Add($"Трасса id {routeId} имеет противоречивые соседние ряды. Трасса сдвинута без изменения формы, но часть шагов может отличаться от заданного.");
        }

        private static void MoveRoutes(Document doc, IReadOnlyList<MepRouteInfo> routes, Dictionary<MepRouteInfo, XYZ> moveVectorsByRoute, SpacingApplyResult result)
        {
            double tolerance = Math.Max(doc.Application.ShortCurveTolerance, MillimetersToInternal(0.5));
            foreach (MepRouteInfo route in routes.Where(route => !route.IsFixed))
            {
                if (!moveVectorsByRoute.TryGetValue(route, out XYZ moveVector))
                    continue;

                if (moveVector.GetLength() < tolerance)
                    continue;

                TryMoveRoute(doc, route, moveVector, result);
            }
        }

        private static void TryMoveRoute(Document doc, MepRouteInfo route, XYZ moveVector, SpacingApplyResult result)
        {
            SubTransaction subTransaction = new SubTransaction(doc);
            try
            {
                subTransaction.Start();
                ElementTransformUtils.MoveElements(doc, route.GetIdsToMove(), moveVector);
                subTransaction.Commit();
                result.MovedElementCount += route.Elements.Count;
            }
            catch (Exception ex)
            {
                try
                {
                    if (subTransaction.GetStatus() == TransactionStatus.Started)
                        subTransaction.RollBack();
                }
                catch
                {
                }

                result.SkippedElementCount += route.Elements.Count;
                result.Messages.Add($"Пропущена трасса id {GetElementIdValue(route.Elements.First().Element.Id)}: {ex.Message}");
            }
        }

        private static SegmentOffset CreateSegmentOffset(Document doc, MepRouteInfo route, MepElementInfo element, XYZ spacingDirection)
        {
            return new SegmentOffset
            {
                Route = route,
                ElementInfo = element,
                CurrentOffset = element.Center.DotProduct(spacingDirection),
                HalfSize = GetHalfSize(doc, element.Element, element.Center, spacingDirection),
                AlignmentOffset = 0
            };
        }

        private static SegmentOffset CreateSegmentOffset(Document doc, MepRouteInfo route, MepElementInfo element, XYZ spacingDirection, XYZ alignmentDirection)
        {
            return new SegmentOffset
            {
                Route = route,
                ElementInfo = element,
                CurrentOffset = element.Center.DotProduct(spacingDirection),
                HalfSize = GetHalfSize(doc, element.Element, element.Center, spacingDirection),
                AlignmentOffset = element.Center.DotProduct(alignmentDirection)
            };
        }

        private static bool IsFixedRoute(Document doc, MepRouteInfo route, HashSet<long> baseElementIdValues)
        {
            if (route.Elements.Any(element => baseElementIdValues.Contains(GetElementIdValue(element.Element.Id)) || element.Element.Pinned))
                return true;

            return route.FittingIds.Any(id => doc.GetElement(id)?.Pinned == true);
        }

        private static List<DirectionFamilyInfo> BuildDirectionFamilies(IReadOnlyList<MepElementInfo> elements)
        {
            List<DirectionFamilyInfo> families = new List<DirectionFamilyInfo>();
            foreach (MepElementInfo element in elements.Where(element => !IsVerticalElement(element)).OrderByDescending(element => element.Length))
            {
                if (families.Any(family => Math.Abs(family.Direction.DotProduct(element.Direction)) >= DirectionTolerance))
                    continue;

                families.Add(new DirectionFamilyInfo
                {
                    Direction = element.Direction,
                    Normal = GetPlanNormal(element.Direction)
                });
            }

            return families;
        }

        private static bool IsVerticalElement(MepElementInfo element)
        {
            return Math.Abs(element.Direction.DotProduct(XYZ.BasisZ)) >= VerticalDirectionTolerance;
        }

        private static List<MepRouteInfo> BuildRoutes(Document doc, List<MepElementInfo> elements)
        {
            Dictionary<long, MepElementInfo> elementById = elements.ToDictionary(element => GetElementIdValue(element.Element.Id));
            Dictionary<long, HashSet<long>> adjacency = elements.ToDictionary(element => GetElementIdValue(element.Element.Id), _ => new HashSet<long>());
            Dictionary<long, HashSet<long>> routeFittingIdsByElementId = elements.ToDictionary(element => GetElementIdValue(element.Element.Id), _ => new HashSet<long>());

            foreach (MepElementInfo element in elements)
            {
                long sourceId = GetElementIdValue(element.Element.Id);
                ConnectorSet sourceConnectors = GetConnectors(element.Element);
                if (sourceConnectors == null)
                    continue;

                foreach (Connector connector in sourceConnectors)
                {
                    foreach (Connector connectedConnector in connector.AllRefs)
                    {
                        Element connectedOwner = connectedConnector.Owner;
                        AddConnection(sourceId, connectedOwner, null, elementById, adjacency, routeFittingIdsByElementId);

                        if (!(connectedOwner is FamilyInstance fitting))
                            continue;

                        ConnectorSet fittingConnectors = GetConnectors(fitting);
                        if (fittingConnectors == null)
                            continue;

                        foreach (Connector fittingConnector in fittingConnectors)
                        {
                            foreach (Connector fittingConnectedConnector in fittingConnector.AllRefs)
                                AddConnection(sourceId, fittingConnectedConnector.Owner, fitting.Id, elementById, adjacency, routeFittingIdsByElementId);
                        }
                    }
                }
            }

            AddEndpointConnections(doc, elements, adjacency);
            HashSet<long> visitedIds = new HashSet<long>();
            List<MepRouteInfo> routes = new List<MepRouteInfo>();

            foreach (MepElementInfo element in elements)
            {
                long id = GetElementIdValue(element.Element.Id);
                if (visitedIds.Contains(id))
                    continue;

                List<MepElementInfo> routeElements = new List<MepElementInfo>();
                HashSet<ElementId> fittingIds = new HashSet<ElementId>(new ElementIdEqualityComparer());
                Queue<long> queue = new Queue<long>();
                queue.Enqueue(id);
                visitedIds.Add(id);

                while (queue.Count > 0)
                {
                    long currentId = queue.Dequeue();
                    MepElementInfo currentElement = elementById[currentId];
                    routeElements.Add(currentElement);

                    foreach (long fittingId in routeFittingIdsByElementId[currentId])
                        fittingIds.Add(ToElementId(fittingId));

                    foreach (long linkedId in adjacency[currentId])
                    {
                        if (visitedIds.Contains(linkedId))
                            continue;

                        visitedIds.Add(linkedId);
                        queue.Enqueue(linkedId);
                    }
                }

                MepRouteInfo route = new MepRouteInfo(routeElements, fittingIds);
                foreach (MepElementInfo routeElement in routeElements)
                    routeElement.Route = route;

                routes.Add(route);
            }

            return routes;
        }

        private static void AddConnection(
            long sourceId,
            Element connectedOwner,
            ElementId fittingId,
            Dictionary<long, MepElementInfo> elementById,
            Dictionary<long, HashSet<long>> adjacency,
            Dictionary<long, HashSet<long>> routeFittingIdsByElementId)
        {
            if (connectedOwner == null)
                return;

            long linkedId = GetElementIdValue(connectedOwner.Id);
            if (linkedId == sourceId || !elementById.ContainsKey(linkedId))
                return;

            adjacency[sourceId].Add(linkedId);
            adjacency[linkedId].Add(sourceId);

            if (fittingId == null)
                return;

            long fittingIdValue = GetElementIdValue(fittingId);
            routeFittingIdsByElementId[sourceId].Add(fittingIdValue);
            routeFittingIdsByElementId[linkedId].Add(fittingIdValue);
        }

        private static void AddEndpointConnections(Document doc, List<MepElementInfo> elements, Dictionary<long, HashSet<long>> adjacency)
        {
            double tolerance = Math.Max(doc.Application.ShortCurveTolerance, MillimetersToInternal(5));
            for (int i = 0; i < elements.Count; i++)
            {
                for (int j = i + 1; j < elements.Count; j++)
                {
                    if (!HasCloseEndpoints(elements[i], elements[j], tolerance))
                        continue;

                    long firstId = GetElementIdValue(elements[i].Element.Id);
                    long secondId = GetElementIdValue(elements[j].Element.Id);
                    adjacency[firstId].Add(secondId);
                    adjacency[secondId].Add(firstId);
                }
            }
        }

        private static bool HasCloseEndpoints(MepElementInfo firstElement, MepElementInfo secondElement, double tolerance)
        {
            return firstElement.StartPoint.DistanceTo(secondElement.StartPoint) <= tolerance
                || firstElement.StartPoint.DistanceTo(secondElement.EndPoint) <= tolerance
                || firstElement.EndPoint.DistanceTo(secondElement.StartPoint) <= tolerance
                || firstElement.EndPoint.DistanceTo(secondElement.EndPoint) <= tolerance;
        }

        private static MepElementInfo CreateInfo(Element element)
        {
            if (!(element.Location is LocationCurve locationCurve) || !(locationCurve.Curve is Line line))
                throw new InvalidOperationException($"Элемент id {GetElementIdValue(element.Id)} не является прямым MEP-участком.");

            if (!IsSupportedMepElement(element))
                throw new InvalidOperationException($"Элемент id {GetElementIdValue(element.Id)} не поддерживается.");

            XYZ startPoint = line.GetEndPoint(0);
            XYZ endPoint = line.GetEndPoint(1);

            return new MepElementInfo
            {
                Element = element,
                StartPoint = startPoint,
                EndPoint = endPoint,
                Center = (startPoint + endPoint).Multiply(0.5),
                Direction = (endPoint - startPoint).Normalize(),
                Length = startPoint.DistanceTo(endPoint)
            };
        }

        private static bool IsSupportedMepElement(Element element)
        {
            return element is Pipe || element is Duct || element is CableTray;
        }

        private static ConnectorSet GetConnectors(Element element)
        {
            if (element is MEPCurve mepCurve)
                return mepCurve.ConnectorManager?.Connectors;

            if (element is FamilyInstance familyInstance)
                return familyInstance.MEPModel?.ConnectorManager?.Connectors;

            return null;
        }

        private static XYZ GetPlanNormal(XYZ direction)
        {
            XYZ normal = direction.CrossProduct(XYZ.BasisZ);
            if (normal.GetLength() < Epsilon)
                normal = direction.CrossProduct(XYZ.BasisX);

            return normal.Normalize();
        }

        private static double GetHalfSize(Document doc, Element element, XYZ center, XYZ spacingDirection)
        {
            double fromGeometry = GetHalfSizeFromGeometry(element, center, spacingDirection);
            if (fromGeometry > Epsilon)
                return fromGeometry;

            double fromBoundingBox = GetHalfSizeFromBoundingBox(element, center, spacingDirection);
            if (fromBoundingBox > Epsilon)
                return fromBoundingBox;

            return GetHalfSizeFromParameters(doc, element);
        }

        private static double GetHalfSizeFromGeometry(Element element, XYZ center, XYZ spacingDirection)
        {
            Options options = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geometry = element.get_Geometry(options);
            if (geometry == null)
                return 0;

            double centerOffset = center.DotProduct(spacingDirection);
            double maxHalfSize = 0;
            foreach (XYZ point in GetGeometryPoints(geometry))
            {
                double halfSize = Math.Abs(point.DotProduct(spacingDirection) - centerOffset);
                if (halfSize > maxHalfSize)
                    maxHalfSize = halfSize;
            }

            return maxHalfSize;
        }

        private static IEnumerable<XYZ> GetGeometryPoints(GeometryElement geometry)
        {
            foreach (GeometryObject geometryObject in geometry)
            {
                if (geometryObject is Solid solid)
                {
                    foreach (Edge edge in solid.Edges)
                    {
                        foreach (XYZ point in edge.Tessellate())
                            yield return point;
                    }
                }
                else if (geometryObject is GeometryInstance instance)
                {
                    foreach (XYZ point in GetGeometryPoints(instance.GetInstanceGeometry()))
                        yield return point;
                }
                else if (geometryObject is Curve curve)
                {
                    yield return curve.GetEndPoint(0);
                    yield return curve.GetEndPoint(1);
                }
            }
        }

        private static double GetHalfSizeFromBoundingBox(Element element, XYZ center, XYZ spacingDirection)
        {
            BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);
            if (boundingBox == null)
                return 0;

            XYZ min = boundingBox.Min;
            XYZ max = boundingBox.Max;
            XYZ[] corners =
            {
                new XYZ(min.X, min.Y, min.Z),
                new XYZ(min.X, min.Y, max.Z),
                new XYZ(min.X, max.Y, min.Z),
                new XYZ(min.X, max.Y, max.Z),
                new XYZ(max.X, min.Y, min.Z),
                new XYZ(max.X, min.Y, max.Z),
                new XYZ(max.X, max.Y, min.Z),
                new XYZ(max.X, max.Y, max.Z)
            };

            double centerOffset = center.DotProduct(spacingDirection);
            return corners.Max(point => Math.Abs(point.DotProduct(spacingDirection) - centerOffset));
        }

        private static double GetHalfSizeFromParameters(Document doc, Element element)
        {
            if (element is Pipe)
                return GetDoubleParam(doc, element, BuiltInParameter.RBS_PIPE_OUTER_DIAMETER) * 0.5;

            double size = 0;
            size = Math.Max(size, GetDoubleParam(doc, element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM));
            size = Math.Max(size, GetDoubleParam(doc, element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM));
            size = Math.Max(size, GetDoubleParam(doc, element, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM));
            size = Math.Max(size, GetDoubleParam(doc, element, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM));

            return size * 0.5;
        }

        private static double GetDoubleParam(Document doc, Element element, BuiltInParameter builtInParameter)
        {
            Parameter parameter = element.get_Parameter(builtInParameter);
            if (parameter != null && parameter.StorageType == StorageType.Double)
                return parameter.AsDouble();

            Element type = doc.GetElement(element.GetTypeId());
            parameter = type?.get_Parameter(builtInParameter);
            if (parameter != null && parameter.StorageType == StorageType.Double)
                return parameter.AsDouble();

            return 0;
        }

        private static double MillimetersToInternal(double value)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
#endif
        }

        private static ElementId ToElementId(long id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return new ElementId((int)id);
#else
            return new ElementId(id);
#endif
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

        private sealed class MepElementInfo
        {
            public Element Element { get; set; }

            public XYZ StartPoint { get; set; }

            public XYZ EndPoint { get; set; }

            public XYZ Center { get; set; }

            public XYZ Direction { get; set; }

            public double Length { get; set; }

            public bool IsFixed { get; set; }

            public MepRouteInfo Route { get; set; }
        }

        private sealed class MepRouteInfo
        {
            private readonly HashSet<ElementId> _fittingIds;

            public MepRouteInfo(IReadOnlyList<MepElementInfo> elements, HashSet<ElementId> fittingIds)
            {
                Elements = elements;
                _fittingIds = fittingIds;
                Center = GetWeightedCenter(elements);
            }

            public IReadOnlyList<MepElementInfo> Elements { get; }

            public IReadOnlyCollection<ElementId> FittingIds => _fittingIds;

            public XYZ Center { get; }

            public bool IsFixed { get; set; }

            public ICollection<ElementId> GetIdsToMove()
            {
                HashSet<ElementId> ids = new HashSet<ElementId>(new ElementIdEqualityComparer());
                foreach (MepElementInfo element in Elements)
                    ids.Add(element.Element.Id);

                foreach (ElementId fittingId in _fittingIds)
                    ids.Add(fittingId);

                return ids;
            }

            private static XYZ GetWeightedCenter(IReadOnlyList<MepElementInfo> elements)
            {
                double summaryLength = elements.Sum(element => element.Length);
                if (summaryLength < Epsilon)
                    return elements.First().Center;

                XYZ summary = XYZ.Zero;
                foreach (MepElementInfo element in elements)
                    summary += element.Center.Multiply(element.Length);

                return summary.Divide(summaryLength);
            }
        }

        private sealed class ElementIdEqualityComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y)
            {
                if (x == null || y == null)
                    return false;

                return GetElementIdValue(x) == GetElementIdValue(y);
            }

            public int GetHashCode(ElementId obj) => GetElementIdValue(obj).GetHashCode();
        }

        private sealed class DirectionFamilyInfo
        {
            public XYZ Direction { get; set; }

            public XYZ Normal { get; set; }
        }

        private sealed class SegmentOffset
        {
            public MepRouteInfo Route { get; set; }

            public MepElementInfo ElementInfo { get; set; }

            public double CurrentOffset { get; set; }

            public double HalfSize { get; set; }

            public double AlignmentOffset { get; set; }
        }

        private sealed class RouteRowOffset
        {
            public MepRouteInfo Route { get; set; }

            public double CurrentOffset { get; set; }

            public double HalfSize { get; set; }
        }

        private sealed class MoveConstraint
        {
            public XYZ Direction { get; set; }

            public double Delta { get; set; }
        }

        private sealed class PairSpacingConstraint
        {
            public MepRouteInfo FirstRoute { get; set; }

            public MepRouteInfo SecondRoute { get; set; }

            public XYZ Direction { get; set; }

            public double Delta { get; set; }
        }

        private sealed class EquationCoefficient
        {
            public int Index { get; set; }

            public double Value { get; set; }
        }
    }

    public sealed class SpacingApplyResult
    {
        public int MovedElementCount { get; set; }

        public int FixedElementCount { get; set; }

        public int SkippedElementCount { get; set; }

        public List<string> Messages { get; } = new List<string>();
    }
}
