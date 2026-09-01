using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TrailingMEP.Common
{
    internal static class RouteBuilder
    {
        public const string PreviewLineStyleName = "KPLN_TRAILING";
        private const double StraightConnectionDotTolerance = 0.999;
        private const double ConnectorPreferenceTolerance = 1.0 / 304.8;

        public static MepCurveData CreateMepCurveData(Document doc, Element element, XYZ targetPoint)
        {
            if (!(element.Location is LocationCurve locationCurve) || !(locationCurve.Curve is Line line))
                throw new InvalidOperationException($"Элемент id {GetElementIdValue(element.Id)} не является прямолинейным участком.");

            XYZ firstEnd = line.GetEndPoint(0);
            XYZ secondEnd = line.GetEndPoint(1);
            XYZ extensionStart = firstEnd.DistanceTo(targetPoint) <= secondEnd.DistanceTo(targetPoint) ? firstEnd : secondEnd;
            XYZ oppositeEnd = extensionStart.IsAlmostEqualTo(firstEnd) ? secondEnd : firstEnd;

            return new MepCurveData
            {
                SourceId = element.Id,
                Kind = GetRouteKind(element),
                TypeId = element.GetTypeId(),
                SystemTypeId = GetSystemTypeId(doc, element),
                LevelId = GetLevelId(element),
                ExtensionStart = extensionStart,
                OppositeEnd = oppositeEnd
            };
        }

        public static IReadOnlyList<ElementId> CreateOrReplacePreviewRoute(
            Document doc,
            IReadOnlyList<ElementId> currentRouteIds,
            XYZ startPoint,
            IReadOnlyList<XYZ> rawRoutePoints,
            XYZ baseDirection,
            bool autoCorrectRoute,
            IReadOnlyList<int> allowedAngles)
        {
            DeletePreviewRoutes(doc, currentRouteIds);

            List<XYZ> routePoints = BuildPreviewRoutePoints(startPoint, rawRoutePoints, baseDirection, autoCorrectRoute, allowedAngles);
            View previewView = GetActivePreviewView(doc);
            GraphicsStyle routeStyle = GetOrCreatePreviewLineStyle(doc);
            List<ElementId> createdIds = new List<ElementId>();

            for (int i = 1; i < routePoints.Count; i++)
            {
                if (routePoints[i - 1].DistanceTo(routePoints[i]) < doc.Application.ShortCurveTolerance)
                    continue;

                Line routeLine = Line.CreateBound(
                    ProjectPointToViewPlane(previewView, routePoints[i - 1]),
                    ProjectPointToViewPlane(previewView, routePoints[i]));
                DetailCurve detailCurve = doc.Create.NewDetailCurve(previewView, routeLine);

                if (routeStyle != null)
                    detailCurve.LineStyle = routeStyle;

                createdIds.Add(detailCurve.Id);
            }

            return createdIds;
        }

        public static void DeletePreviewRoutes(Document doc, IReadOnlyList<ElementId> routeIds)
        {
            if (routeIds == null)
                return;

            foreach (ElementId routeId in routeIds)
            {
                if (routeId != null && !routeId.Equals(ElementId.InvalidElementId) && doc.GetElement(routeId) != null)
                    doc.Delete(routeId);
            }
        }

        public static IReadOnlyList<ElementId> BuildExtensions(Document doc, IReadOnlyList<MepCurveData> sourceCurves, IReadOnlyList<XYZ> baseRoutePoints)
        {
            if (baseRoutePoints == null || baseRoutePoints.Count < 2)
                return new List<ElementId>();

            IReadOnlyList<XYZ> routePoints = MergeCollinearRoutePoints(baseRoutePoints);
            if (routePoints.Count < 2)
                return new List<ElementId>();

            List<ElementId> createdIds = new List<ElementId>();

            foreach (MepCurveData sourceData in sourceCurves)
            {
                Element sourceElement = doc.GetElement(sourceData.SourceId);
                IReadOnlyList<XYZ> elementRoutePoints = BuildElementOffsetRoutePoints(routePoints, sourceData);
                List<Element> createdRouteElements = new List<Element>();
                int firstRoutePointIndex = 1;
                XYZ sourceConnectionPoint = sourceData.ExtensionStart;

                if (elementRoutePoints.Count > 1 && TryExtendSourceElement(sourceElement, sourceData, elementRoutePoints[1]))
                {
                    createdIds.Add(sourceElement.Id);
                    firstRoutePointIndex = 2;
                    sourceConnectionPoint = elementRoutePoints[1];
                    doc.Regenerate();
                }

                for (int i = firstRoutePointIndex; i < elementRoutePoints.Count; i++)
                {
                    XYZ startPoint = elementRoutePoints[i - 1];
                    XYZ endPoint = elementRoutePoints[i];

                    if (startPoint.DistanceTo(endPoint) < doc.Application.ShortCurveTolerance)
                        continue;

                    Element created = CreateExtension(doc, sourceData, startPoint, endPoint);
                    CopySizeParameters(sourceElement, created);

                    if (created == null)
                        continue;

                    createdRouteElements.Add(created);
                    createdIds.Add(created.Id);
                }

                doc.Regenerate();
                ConnectRouteElements(doc, sourceElement, createdRouteElements, sourceConnectionPoint);
            }

            return createdIds;
        }

        private static IReadOnlyList<XYZ> MergeCollinearRoutePoints(IReadOnlyList<XYZ> routePoints)
        {
            List<XYZ> result = new List<XYZ>();
            if (routePoints == null || routePoints.Count == 0)
                return result;

            result.Add(routePoints[0]);

            for (int i = 1; i < routePoints.Count - 1; i++)
            {
                XYZ previousDirection = GetSegmentDirectionXY(result.Last(), routePoints[i]);
                XYZ nextDirection = GetSegmentDirectionXY(routePoints[i], routePoints[i + 1]);

                if (previousDirection.DotProduct(nextDirection) >= StraightConnectionDotTolerance)
                    continue;

                result.Add(routePoints[i]);
            }

            result.Add(routePoints.Last());
            return result;
        }

        public static IReadOnlyList<XYZ> GetPreviewRoutePoints(Document doc, IReadOnlyList<ElementId> previewRouteIds, XYZ basePoint)
        {
            List<XYZ> routePoints = new List<XYZ>();
            if (basePoint == null || previewRouteIds == null || previewRouteIds.Count == 0)
                return routePoints;

            routePoints.Add(basePoint);
            List<Line> routeLines = new List<Line>();

            foreach (ElementId previewRouteId in previewRouteIds)
            {
                if (previewRouteId == null || previewRouteId.Equals(ElementId.InvalidElementId))
                    continue;

                if (!(doc.GetElement(previewRouteId) is CurveElement curveElement) || !(curveElement.GeometryCurve is Line line))
                    continue;

                routeLines.Add(line);
            }

            List<XYZ> orderedPoints = BuildOrderedRouteLinePoints(routeLines, basePoint);
            if (orderedPoints.Count < 2)
                return routePoints;

            XYZ routeStartPoint = orderedPoints.First();
            XYZ startShift = new XYZ(basePoint.X - routeStartPoint.X, basePoint.Y - routeStartPoint.Y, 0);

            for (int i = 1; i < orderedPoints.Count; i++)
            {
                XYZ shiftedPoint = orderedPoints[i] + startShift;
                XYZ nextPoint = new XYZ(shiftedPoint.X, shiftedPoint.Y, basePoint.Z);

                if (DistanceXY(routePoints.Last(), nextPoint) < doc.Application.ShortCurveTolerance)
                    continue;

                routePoints.Add(nextPoint);
            }

            return routePoints;
        }

        private static List<XYZ> BuildOrderedRouteLinePoints(List<Line> routeLines, XYZ basePoint)
        {
            List<XYZ> orderedPoints = new List<XYZ>();
            if (routeLines == null || routeLines.Count == 0)
                return orderedPoints;

            int startLineIndex = -1;
            XYZ startPoint = null;
            XYZ currentPoint = null;
            double nearestStartDistance = double.MaxValue;

            for (int i = 0; i < routeLines.Count; i++)
            {
                Line line = routeLines[i];
                XYZ firstEnd = line.GetEndPoint(0);
                XYZ secondEnd = line.GetEndPoint(1);
                double firstDistance = DistanceXY(basePoint, firstEnd);
                double secondDistance = DistanceXY(basePoint, secondEnd);

                if (firstDistance < nearestStartDistance)
                {
                    startLineIndex = i;
                    nearestStartDistance = firstDistance;
                    startPoint = firstEnd;
                    currentPoint = secondEnd;
                }

                if (secondDistance < nearestStartDistance)
                {
                    startLineIndex = i;
                    nearestStartDistance = secondDistance;
                    startPoint = secondEnd;
                    currentPoint = firstEnd;
                }
            }

            if (startLineIndex < 0 || startPoint == null || currentPoint == null)
                return orderedPoints;

            orderedPoints.Add(startPoint);
            orderedPoints.Add(currentPoint);
            routeLines.RemoveAt(startLineIndex);

            while (routeLines.Count > 0)
            {
                int nearestLineIndex = -1;
                XYZ nextPoint = null;
                double nearestDistance = double.MaxValue;

                for (int i = 0; i < routeLines.Count; i++)
                {
                    Line line = routeLines[i];
                    XYZ firstEnd = line.GetEndPoint(0);
                    XYZ secondEnd = line.GetEndPoint(1);
                    double firstDistance = DistanceXY(currentPoint, firstEnd);
                    double secondDistance = DistanceXY(currentPoint, secondEnd);
                    double lineDistance = Math.Min(firstDistance, secondDistance);

                    if (lineDistance >= nearestDistance)
                        continue;

                    nearestLineIndex = i;
                    nearestDistance = lineDistance;
                    nextPoint = firstDistance <= secondDistance ? secondEnd : firstEnd;
                }

                if (nearestLineIndex < 0 || nextPoint == null)
                    break;

                currentPoint = nextPoint;
                orderedPoints.Add(currentPoint);
                routeLines.RemoveAt(nearestLineIndex);
            }

            return orderedPoints;
        }

        public static List<XYZ> BuildPreviewRoutePoints(
            XYZ startPoint,
            IReadOnlyList<XYZ> rawRoutePoints,
            XYZ baseDirection,
            bool autoCorrectRoute,
            IReadOnlyList<int> allowedAngles)
        {
            List<XYZ> routePoints = new List<XYZ>();
            if (startPoint == null)
                return routePoints;

            routePoints.Add(startPoint);
            if (rawRoutePoints == null || rawRoutePoints.Count == 0)
                return routePoints;

            XYZ previousDirection = ProjectToXY(baseDirection);
            if (previousDirection.GetLength() < 1e-9)
                previousDirection = XYZ.BasisX;

            previousDirection = previousDirection.Normalize();

            foreach (XYZ rawPoint in rawRoutePoints)
            {
                XYZ previousPoint = routePoints.Last();
                XYZ targetPoint = new XYZ(rawPoint.X, rawPoint.Y, startPoint.Z);
                XYZ desiredVector = ProjectToXY(targetPoint - previousPoint);
                double length = desiredVector.GetLength();

                if (length < 1e-9)
                    continue;

                XYZ routeDirection = desiredVector.Normalize();
                if (autoCorrectRoute)
                    routeDirection = GetNearestAllowedDirection(previousDirection, routeDirection, allowedAngles);

                XYZ nextPoint = previousPoint + routeDirection.Multiply(length);
                routePoints.Add(new XYZ(nextPoint.X, nextPoint.Y, startPoint.Z));
                previousDirection = routeDirection;
            }

            return routePoints;
        }

        public static XYZ ProjectToXY(XYZ vector)
        {
            if (vector == null)
                return XYZ.Zero;

            return new XYZ(vector.X, vector.Y, 0);
        }

        public static double DistanceXY(XYZ firstPoint, XYZ secondPoint)
        {
            if (firstPoint == null || secondPoint == null)
                return double.MaxValue;

            return ProjectToXY(firstPoint - secondPoint).GetLength();
        }

        private static Element CreateExtension(Document doc, MepCurveData sourceData, XYZ startPoint, XYZ endPoint)
        {
            switch (sourceData.Kind)
            {
                case MepRouteKind.Pipe:
                    return Pipe.Create(doc, sourceData.SystemTypeId, sourceData.TypeId, sourceData.LevelId, startPoint, endPoint);
                case MepRouteKind.Duct:
                    return Duct.Create(doc, sourceData.SystemTypeId, sourceData.TypeId, sourceData.LevelId, startPoint, endPoint);
                case MepRouteKind.CableTray:
                    return CableTray.Create(doc, sourceData.TypeId, startPoint, endPoint, sourceData.LevelId);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static IReadOnlyList<XYZ> BuildElementOffsetRoutePoints(IReadOnlyList<XYZ> baseRoutePoints, MepCurveData sourceData)
        {
            List<XYZ> result = new List<XYZ>();
            if (baseRoutePoints == null || baseRoutePoints.Count == 0)
                return result;

            result.Add(sourceData.ExtensionStart);
            if (baseRoutePoints.Count == 1)
                return result;

            XYZ firstDirection = GetSegmentDirectionXY(baseRoutePoints[0], baseRoutePoints[1]);
            XYZ firstNormal = GetLeftNormalXY(firstDirection);
            XYZ startOffset = ProjectToXY(sourceData.ExtensionStart - baseRoutePoints[0]);
            double lateralOffset = startOffset.DotProduct(firstNormal);

            for (int i = 1; i < baseRoutePoints.Count - 1; i++)
            {
                XYZ previousDirection = GetSegmentDirectionXY(baseRoutePoints[i - 1], baseRoutePoints[i]);
                XYZ nextDirection = GetSegmentDirectionXY(baseRoutePoints[i], baseRoutePoints[i + 1]);

                XYZ previousPoint = OffsetPointXY(baseRoutePoints[i], previousDirection, lateralOffset, sourceData.ExtensionStart.Z);
                XYZ nextPoint = OffsetPointXY(baseRoutePoints[i], nextDirection, lateralOffset, sourceData.ExtensionStart.Z);
                XYZ intersection = TryIntersectLinesXY(previousPoint, previousDirection, nextPoint, nextDirection);

                result.Add(intersection ?? nextPoint);
            }

            XYZ lastDirection = GetSegmentDirectionXY(baseRoutePoints[baseRoutePoints.Count - 2], baseRoutePoints[baseRoutePoints.Count - 1]);
            result.Add(OffsetPointXY(baseRoutePoints.Last(), lastDirection, lateralOffset, sourceData.ExtensionStart.Z));

            return result;
        }

        private static XYZ OffsetPointXY(XYZ basePoint, XYZ direction, double lateralOffset, double z)
        {
            XYZ normal = GetLeftNormalXY(direction);
            return new XYZ(basePoint.X + normal.X * lateralOffset, basePoint.Y + normal.Y * lateralOffset, z);
        }

        private static XYZ GetSegmentDirectionXY(XYZ startPoint, XYZ endPoint)
        {
            XYZ direction = ProjectToXY(endPoint - startPoint);
            return direction.GetLength() > 1e-9 ? direction.Normalize() : XYZ.BasisX;
        }

        private static XYZ GetLeftNormalXY(XYZ direction)
        {
            XYZ normalized = ProjectToXY(direction);
            if (normalized.GetLength() < 1e-9)
                normalized = XYZ.BasisX;

            normalized = normalized.Normalize();
            return new XYZ(-normalized.Y, normalized.X, 0);
        }

        private static XYZ TryIntersectLinesXY(XYZ firstPoint, XYZ firstDirection, XYZ secondPoint, XYZ secondDirection)
        {
            double cross = CrossXY(firstDirection, secondDirection);
            if (Math.Abs(cross) < 1e-9)
                return null;

            XYZ delta = secondPoint - firstPoint;
            double firstParameter = CrossXY(ProjectToXY(delta), secondDirection) / cross;
            XYZ point = firstPoint + firstDirection.Multiply(firstParameter);
            return new XYZ(point.X, point.Y, firstPoint.Z);
        }

        private static double CrossXY(XYZ firstVector, XYZ secondVector)
        {
            return firstVector.X * secondVector.Y - firstVector.Y * secondVector.X;
        }

        private static XYZ GetNearestAllowedDirection(XYZ previousDirection, XYZ desiredDirection, IReadOnlyList<int> allowedAngles)
        {
            List<XYZ> candidates = new List<XYZ>
            {
                previousDirection.Normalize()
            };

            foreach (int angle in allowedAngles ?? new List<int>())
            {
                AddAngleCandidates(candidates, previousDirection, angle);

                if (angle != 90)
                    AddAngleCandidates(candidates, previousDirection, 180 - angle);
            }

            return candidates
                .OrderByDescending(candidate => candidate.Normalize().DotProduct(desiredDirection))
                .First()
                .Normalize();
        }

        private static void AddAngleCandidates(List<XYZ> candidates, XYZ direction, int angle)
        {
            candidates.Add(RotateXY(direction, angle));
            candidates.Add(RotateXY(direction, -angle));
        }

        private static XYZ RotateXY(XYZ direction, double angle)
        {
            double radians = angle * Math.PI / 180.0;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);

            return new XYZ(
                direction.X * cos - direction.Y * sin,
                direction.X * sin + direction.Y * cos,
                0).Normalize();
        }

        private static bool TryExtendSourceElement(Element sourceElement, MepCurveData sourceData, XYZ newEndPoint)
        {
            if (!(sourceElement?.Location is LocationCurve locationCurve) || !(locationCurve.Curve is Line sourceLine))
                return false;

            XYZ extensionDirection = GetSegmentDirectionXY(sourceData.ExtensionStart, newEndPoint);
            XYZ sourceDirection = ProjectToXY(sourceData.SourceDirection);

            if (sourceDirection.GetLength() < 1e-9 || extensionDirection.GetLength() < 1e-9)
                return false;

            if (sourceDirection.Normalize().DotProduct(extensionDirection.Normalize()) < StraightConnectionDotTolerance)
                return false;

            XYZ firstEnd = sourceLine.GetEndPoint(0);
            XYZ secondEnd = sourceLine.GetEndPoint(1);
            if (firstEnd.DistanceTo(sourceData.ExtensionStart) <= secondEnd.DistanceTo(sourceData.ExtensionStart))
                locationCurve.Curve = Line.CreateBound(newEndPoint, secondEnd);
            else
                locationCurve.Curve = Line.CreateBound(firstEnd, newEndPoint);

            return true;
        }

        private static void ConnectRouteElements(Document doc, Element sourceElement, List<Element> createdRouteElements, XYZ sourceConnectionPoint)
        {
            if (sourceElement == null || createdRouteElements == null || createdRouteElements.Count == 0)
                return;

            TryConnectElements(doc, sourceElement, createdRouteElements.First(), sourceConnectionPoint, false);
            doc.Regenerate();

            for (int i = 1; i < createdRouteElements.Count; i++)
            {
                Connector connection = GetNearestConnector(createdRouteElements[i - 1], GetCurveEndPoint(createdRouteElements[i - 1], 1));
                XYZ connectionPoint = connection?.Origin ?? GetCurveEndPoint(createdRouteElements[i - 1], 1);
                TryConnectElements(doc, createdRouteElements[i - 1], createdRouteElements[i], connectionPoint, true);
                doc.Regenerate();
            }
        }

        private static bool TryConnectElements(Document doc, Element firstElement, Element secondElement, XYZ connectionPoint, bool allowUnionFitting)
        {
            Connector firstConnector = GetNearestConnector(firstElement, connectionPoint);
            Connector secondConnector = GetNearestConnector(secondElement, connectionPoint);

            if (firstConnector == null || secondConnector == null)
                return false;

            if (firstConnector.IsConnectedTo(secondConnector))
                return true;

            bool isStraightConnection = ShouldCreateUnionFitting(firstElement, secondElement, connectionPoint);
            if (TryCreateConnectionFitting(
                doc,
                firstElement,
                secondElement,
                firstConnector,
                secondConnector,
                connectionPoint,
                allowUnionFitting,
                isStraightConnection))
                return true;

            try
            {
                firstConnector.ConnectTo(secondConnector);
                if (firstConnector.IsConnectedTo(secondConnector))
                    return true;
            }
            catch
            {
            }

            return false;
        }

        private static bool TryCreateConnectionFitting(
            Document doc,
            Element firstElement,
            Element secondElement,
            Connector firstConnector,
            Connector secondConnector,
            XYZ connectionPoint,
            bool allowUnionFitting,
            bool isStraightConnection)
        {
            if (isStraightConnection)
                return allowUnionFitting && TryCreateUnionFitting(doc, firstConnector, secondConnector);

            try
            {
                doc.Create.NewElbowFitting(firstConnector, secondConnector);
                doc.Regenerate();
                return true;
            }
            catch
            {
            }

            return false;
        }

        private static bool TryCreateUnionFitting(Document doc, Connector firstConnector, Connector secondConnector)
        {
            try
            {
                doc.Create.NewUnionFitting(firstConnector, secondConnector);
                doc.Regenerate();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Connector GetNearestConnector(Element element, XYZ point)
        {
            ConnectorSet connectors = GetConnectors(element);
            if (connectors == null)
                return null;

            Connector nearestConnector = null;
            Connector nearestOpenConnector = null;
            double nearestDistance = double.MaxValue;
            double nearestOpenDistance = double.MaxValue;

            foreach (Connector connector in connectors)
            {
                if (connector.ConnectorType != ConnectorType.End)
                    continue;

                double distance = connector.Origin.DistanceTo(point);
                if (distance >= nearestDistance)
                {
                    if (!connector.IsConnected && distance < nearestOpenDistance)
                    {
                        nearestOpenConnector = connector;
                        nearestOpenDistance = distance;
                    }

                    continue;
                }

                nearestConnector = connector;
                nearestDistance = distance;

                if (!connector.IsConnected)
                {
                    nearestOpenConnector = connector;
                    nearestOpenDistance = distance;
                }
            }

            if (nearestOpenConnector != null && nearestOpenDistance <= nearestDistance + ConnectorPreferenceTolerance)
                return nearestOpenConnector;

            return nearestConnector;
        }

        private static ConnectorSet GetConnectors(Element element)
        {
            if (element is MEPCurve mepCurve)
                return mepCurve.ConnectorManager?.Connectors;

            if (element is FamilyInstance familyInstance)
                return familyInstance.MEPModel?.ConnectorManager?.Connectors;

            return null;
        }

        private static XYZ GetCurveEndPoint(Element element, int index)
        {
            if (element?.Location is LocationCurve locationCurve)
                return locationCurve.Curve.GetEndPoint(index);

            return XYZ.Zero;
        }

        private static bool ShouldCreateUnionFitting(Element firstElement, Element secondElement, XYZ connectionPoint)
        {
            XYZ firstDirection = GetDirectionFromConnectionPoint(firstElement, connectionPoint);
            XYZ secondDirection = GetDirectionFromConnectionPoint(secondElement, connectionPoint);

            if (firstDirection == null || secondDirection == null)
                return false;

            return Math.Abs(firstDirection.Normalize().DotProduct(secondDirection.Normalize())) >= StraightConnectionDotTolerance;
        }

        private static XYZ GetDirectionFromConnectionPoint(Element element, XYZ connectionPoint)
        {
            if (!(element?.Location is LocationCurve locationCurve) || !(locationCurve.Curve is Line line))
                return null;

            XYZ firstEnd = line.GetEndPoint(0);
            XYZ secondEnd = line.GetEndPoint(1);

            return firstEnd.DistanceTo(connectionPoint) <= secondEnd.DistanceTo(connectionPoint)
                ? secondEnd - firstEnd
                : firstEnd - secondEnd;
        }

        private static GraphicsStyle GetOrCreatePreviewLineStyle(Document doc)
        {
            Category linesCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            Category previewCategory = null;

            foreach (Category subCategory in linesCategory.SubCategories)
            {
                if (subCategory.Name == PreviewLineStyleName)
                {
                    previewCategory = subCategory;
                    break;
                }
            }

            if (previewCategory == null)
            {
                previewCategory = doc.Settings.Categories.NewSubcategory(linesCategory, PreviewLineStyleName);
                previewCategory.LineColor = new Color(255, 128, 0);
                previewCategory.SetLineWeight(6, GraphicsStyleType.Projection);
            }

            return previewCategory.GetGraphicsStyle(GraphicsStyleType.Projection);
        }

        private static View GetActivePreviewView(Document doc)
        {
            View view = doc.ActiveView;
            if (view == null || view.IsTemplate || view.ViewType == ViewType.ThreeD)
                throw new InvalidOperationException("Линию детализации нельзя создать на текущем виде. Открой план, разрез или фасад.");

            return view;
        }

        private static XYZ ProjectPointToViewPlane(View view, XYZ point)
        {
            XYZ normal = view.ViewDirection.Normalize();
            XYZ origin = view.Origin;
            double offset = (point - origin).DotProduct(normal);
            return point - normal.Multiply(offset);
        }

        private static MepRouteKind GetRouteKind(Element element)
        {
            if (element is Pipe)
                return MepRouteKind.Pipe;

            if (element is Duct)
                return MepRouteKind.Duct;

            if (element is CableTray)
                return MepRouteKind.CableTray;

            throw new InvalidOperationException($"Элемент id {GetElementIdValue(element.Id)} не поддерживается.");
        }

        private static ElementId GetSystemTypeId(Document doc, Element element)
        {
            ElementId systemTypeId = ElementId.InvalidElementId;

            if (element is Pipe pipe && pipe.MEPSystem != null)
                systemTypeId = pipe.MEPSystem.GetTypeId();
            else if (element is Duct duct && duct.MEPSystem != null)
                systemTypeId = duct.MEPSystem.GetTypeId();

            if (IsValidElementId(systemTypeId))
                return systemTypeId;

            Parameter pipeSystemParam = element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
            if (pipeSystemParam != null && IsValidElementId(pipeSystemParam.AsElementId()))
                return pipeSystemParam.AsElementId();

            Parameter ductSystemParam = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
            if (ductSystemParam != null && IsValidElementId(ductSystemParam.AsElementId()))
                return ductSystemParam.AsElementId();

            if (element is Pipe)
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(PipingSystemType))
                    .FirstElementId();

            if (element is Duct)
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(DuctSystemType))
                    .FirstElementId();

            return ElementId.InvalidElementId;
        }

        private static ElementId GetLevelId(Element element)
        {
            Parameter levelParam = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
            if (levelParam != null && IsValidElementId(levelParam.AsElementId()))
                return levelParam.AsElementId();

            return element.LevelId;
        }

        private static void CopySizeParameters(Element source, Element target)
        {
            if (source == null || target == null)
                return;

            CopyDoubleParam(source, target, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            CopyDoubleParam(source, target, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            CopyDoubleParam(source, target, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            CopyDoubleParam(source, target, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            CopyDoubleParam(source, target, BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            CopyDoubleParam(source, target, BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
        }

        private static void CopyDoubleParam(Element source, Element target, BuiltInParameter builtInParameter)
        {
            Parameter sourceParam = source.get_Parameter(builtInParameter);
            Parameter targetParam = target.get_Parameter(builtInParameter);

            if (sourceParam == null || targetParam == null || targetParam.IsReadOnly)
                return;

            if (sourceParam.StorageType == StorageType.Double && targetParam.StorageType == StorageType.Double)
                targetParam.Set(sourceParam.AsDouble());
        }

        private static bool IsValidElementId(ElementId id)
        {
            return id != null && !id.Equals(ElementId.InvalidElementId);
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }
    }
}
