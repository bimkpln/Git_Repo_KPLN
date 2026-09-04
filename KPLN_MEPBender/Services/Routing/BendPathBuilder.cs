using Autodesk.Revit.DB;
using KPLN_MEPBender.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class BendPathBuilder
    {
        private static readonly double MinSegmentLength = UnitConvert.MmToInternal(10);
        private readonly ObstacleOutlineBuilder _obstacleOutlineBuilder = new ObstacleOutlineBuilder();

        public bool TryBuild(MepBendRequest request, MEPCurve source, Line routeLine, Outline obstacleOutline, out List<XYZ> points)
        {
            points = null;

            XYZ start = routeLine.GetEndPoint(0);
            XYZ end = routeLine.GetEndPoint(1);
            XYZ routeDirection = (end - start).Normalize();
            double routeLength = start.DistanceTo(end);

            double clearance = UnitConvert.MmToInternal(request.OffsetMm);

            if (!RouteMayIntersectObstacle(source, obstacleOutline, clearance))
                return false;

            ProjectionRange range = GetProjectionRange(start, routeDirection, obstacleOutline);
            if (range.Max < MinSegmentLength || range.Min > routeLength - MinSegmentLength)
                return false;

            double angleRadians = request.AngleDegrees * Math.PI / 180.0;

            foreach (BendDirection direction in request.Directions)
            {
                XYZ offsetDirection = ResolveOffsetDirection(direction, request.ActiveView, routeDirection);
                if (offsetDirection == null)
                    continue;

                double centerOffset = CalculateCenterOffset(request, source, start, routeDirection, obstacleOutline, offsetDirection, direction, clearance);
                if (centerOffset <= MinSegmentLength)
                    continue;

                double run = Math.Abs(request.AngleDegrees - 90) < 0.001
                    ? 0
                    : centerOffset / Math.Tan(angleRadians);

                double lateralStagger = GetLateralLongitudinalStagger(request, source, routeDirection, offsetDirection, direction);
                double bendStartT = range.Min - clearance - run - lateralStagger;
                double bendEndT = range.Max + clearance + run + lateralStagger;

                if (bendStartT < MinSegmentLength || bendEndT > routeLength - MinSegmentLength)
                    continue;

                XYZ bendStart = start + routeDirection * bendStartT;
                XYZ bendEnd = start + routeDirection * bendEndT;

                List<XYZ> candidate = new List<XYZ>
                {
                    start,
                    bendStart,
                    bendStart + routeDirection * run + offsetDirection * centerOffset,
                    bendEnd - routeDirection * run + offsetDirection * centerOffset,
                    bendEnd,
                    end
                };

                points = RemoveShortSegments(candidate);
                if (points.Count >= 4 && HasValidSegments(points))
                    return true;
            }

            return false;
        }

        private bool RouteMayIntersectObstacle(MEPCurve source, Outline obstacleOutline, double clearance)
        {
            BoundingBoxXYZ sourceBox = source.get_BoundingBox(null);
            if (sourceBox == null)
                return false;

            Outline routeOutline = new Outline(sourceBox.Min, sourceBox.Max);

            return routeOutline.Intersects(obstacleOutline, clearance);
        }

        private ProjectionRange GetProjectionRange(XYZ start, XYZ routeDirection, Outline outline)
        {
            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (XYZ point in GetOutlineCorners(outline))
            {
                double projection = (point - start).DotProduct(routeDirection);
                min = Math.Min(min, projection);
                max = Math.Max(max, projection);
            }

            return new ProjectionRange(min, max);
        }

        private double CalculateCenterOffset(
            MepBendRequest request,
            MEPCurve source,
            XYZ routeStart,
            XYZ routeDirection,
            Outline obstacleOutline,
            XYZ offsetDirection,
            BendDirection bendDirection,
            double clearance)
        {
            if (bendDirection == BendDirection.Left || bendDirection == BendDirection.Right)
                return CalculateLateralGroupOffset(request, source, routeDirection, obstacleOutline, offsetDirection, clearance);

            if (bendDirection == BendDirection.Down && request.AlignVerticalBendByLowest)
                return CalculateLowestAlignedVerticalOffset(request, source, routeStart, routeDirection, obstacleOutline, offsetDirection, clearance);

            return CalculateSingleRouteOffset(source, routeStart, obstacleOutline, offsetDirection, clearance);
        }

        private double CalculateSingleRouteOffset(MEPCurve source, XYZ routeStart, Outline obstacleOutline, XYZ offsetDirection, double clearance)
        {
            double sourceCenterProjection = routeStart.DotProduct(offsetDirection);
            return CalculateTargetCenterProjection(source, routeStart, obstacleOutline, offsetDirection, clearance) - sourceCenterProjection;
        }

        private double CalculateLowestAlignedVerticalOffset(
            MepBendRequest request,
            MEPCurve source,
            XYZ routeStart,
            XYZ routeDirection,
            Outline obstacleOutline,
            XYZ offsetDirection,
            double clearance)
        {
            double sourceCenterProjection = routeStart.DotProduct(offsetDirection);
            double targetCenterProjection = CalculateTargetCenterProjection(source, routeStart, obstacleOutline, offsetDirection, clearance);

            foreach (ElementId routeElementId in request.RouteElementIds)
            {
                MEPCurve route = request.Doc.GetElement(routeElementId) as MEPCurve;
                if (route == null || route.Id == source.Id)
                    continue;

                LocationCurve locationCurve = route.Location as LocationCurve;
                Line routeLine = locationCurve?.Curve as Line;
                if (routeLine == null)
                    continue;

                if (!IsParallelTo(route, routeDirection))
                    continue;

                Outline routeObstacleOutline = _obstacleOutlineBuilder.BuildForRoute(request, route, 0, 0);
                if (routeObstacleOutline == null)
                    continue;

                double routeTargetProjection = CalculateTargetCenterProjection(route, routeLine.GetEndPoint(0), routeObstacleOutline, offsetDirection, clearance);
                targetCenterProjection = Math.Max(targetCenterProjection, routeTargetProjection);
            }

            return targetCenterProjection - sourceCenterProjection;
        }

        private double CalculateTargetCenterProjection(MEPCurve source, XYZ routeStart, Outline obstacleOutline, XYZ offsetDirection, double clearance)
        {
            double sourceCenterProjection = routeStart.DotProduct(offsetDirection);
            double obstacleSurfaceProjection = GetMaxProjection(obstacleOutline, offsetDirection);
            double routeHalfSize = GetRouteHalfSize(source, offsetDirection, sourceCenterProjection);

            return obstacleSurfaceProjection + clearance + routeHalfSize;
        }

        private double CalculateLateralGroupOffset(
            MepBendRequest request,
            MEPCurve source,
            XYZ routeDirection,
            Outline obstacleOutline,
            XYZ offsetDirection,
            double clearance)
        {
            double groupNearSurfaceProjection = GetNearSurfaceProjection(source, offsetDirection);

            foreach (ElementId routeElementId in request.RouteElementIds)
            {
                MEPCurve route = request.Doc.GetElement(routeElementId) as MEPCurve;
                if (route == null || route.Id == source.Id)
                    continue;

                if (!IsParallelTo(route, routeDirection))
                    continue;

                groupNearSurfaceProjection = Math.Min(groupNearSurfaceProjection, GetNearSurfaceProjection(route, offsetDirection));
            }

            return GetMaxProjection(obstacleOutline, offsetDirection) + clearance - groupNearSurfaceProjection;
        }

        private double GetLateralLongitudinalStagger(
            MepBendRequest request,
            MEPCurve source,
            XYZ routeDirection,
            XYZ offsetDirection,
            BendDirection bendDirection)
        {
            if (bendDirection != BendDirection.Left && bendDirection != BendDirection.Right)
                return 0;

            double sourceCenterProjection = GetCenterProjection(source, offsetDirection);
            double minCenterProjection = sourceCenterProjection;

            foreach (ElementId routeElementId in request.RouteElementIds)
            {
                MEPCurve route = request.Doc.GetElement(routeElementId) as MEPCurve;
                if (route == null)
                    continue;

                if (!IsParallelTo(route, routeDirection))
                    continue;

                minCenterProjection = Math.Min(minCenterProjection, GetCenterProjection(route, offsetDirection));
            }

            return Math.Abs(sourceCenterProjection - minCenterProjection);
        }

        private double GetRouteHalfSize(MEPCurve source, XYZ offsetDirection, double centerProjection)
        {
            BoundingBoxXYZ sourceBox = source.get_BoundingBox(null);
            if (sourceBox == null)
                return 0;

            double maxProjection = GetMaxProjection(new Outline(sourceBox.Min, sourceBox.Max), offsetDirection);
            return Math.Max(0, maxProjection - centerProjection);
        }

        private double GetNearSurfaceProjection(MEPCurve source, XYZ offsetDirection)
        {
            BoundingBoxXYZ sourceBox = source.get_BoundingBox(null);
            if (sourceBox == null)
            {
                LocationCurve locationCurve = source.Location as LocationCurve;
                Line line = locationCurve?.Curve as Line;
                if (line != null)
                    return line.GetEndPoint(0).DotProduct(offsetDirection);

                return 0;
            }

            return GetMinProjection(new Outline(sourceBox.Min, sourceBox.Max), offsetDirection);
        }

        private double GetCenterProjection(MEPCurve route, XYZ offsetDirection)
        {
            LocationCurve locationCurve = route.Location as LocationCurve;
            Line line = locationCurve?.Curve as Line;
            if (line == null)
                return 0;

            XYZ center = (line.GetEndPoint(0) + line.GetEndPoint(1)) * 0.5;
            return center.DotProduct(offsetDirection);
        }

        private double GetMaxProjection(Outline outline, XYZ direction)
        {
            double maxProjection = double.MinValue;

            foreach (XYZ point in GetOutlineCorners(outline))
                maxProjection = Math.Max(maxProjection, point.DotProduct(direction));

            return maxProjection;
        }

        private double GetMinProjection(Outline outline, XYZ direction)
        {
            double minProjection = double.MaxValue;

            foreach (XYZ point in GetOutlineCorners(outline))
                minProjection = Math.Min(minProjection, point.DotProduct(direction));

            return minProjection;
        }

        private bool IsParallelTo(MEPCurve route, XYZ routeDirection)
        {
            LocationCurve locationCurve = route.Location as LocationCurve;
            Line line = locationCurve?.Curve as Line;
            if (line == null)
                return false;

            XYZ candidateDirection = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
            return Math.Abs(Math.Abs(candidateDirection.DotProduct(routeDirection)) - 1) < 0.01;
        }

        private IEnumerable<XYZ> GetOutlineCorners(Outline outline)
        {
            XYZ min = outline.MinimumPoint;
            XYZ max = outline.MaximumPoint;

            yield return new XYZ(min.X, min.Y, min.Z);
            yield return new XYZ(min.X, min.Y, max.Z);
            yield return new XYZ(min.X, max.Y, min.Z);
            yield return new XYZ(min.X, max.Y, max.Z);
            yield return new XYZ(max.X, min.Y, min.Z);
            yield return new XYZ(max.X, min.Y, max.Z);
            yield return new XYZ(max.X, max.Y, min.Z);
            yield return new XYZ(max.X, max.Y, max.Z);
        }

        private XYZ ResolveOffsetDirection(BendDirection direction, View activeView, XYZ routeDirection)
        {
            XYZ candidate;
            switch (direction)
            {
                case BendDirection.Up:
                    candidate = XYZ.BasisZ;
                    break;
                case BendDirection.Down:
                    candidate = -XYZ.BasisZ;
                    break;
                case BendDirection.Left:
                    candidate = -(activeView?.RightDirection ?? XYZ.BasisX);
                    break;
                case BendDirection.Right:
                    candidate = activeView?.RightDirection ?? XYZ.BasisX;
                    break;
                default:
                    return null;
            }

            XYZ perpendicular = candidate - routeDirection * candidate.DotProduct(routeDirection);
            if (perpendicular.GetLength() < 0.001)
                return null;

            return perpendicular.Normalize();
        }

        private List<XYZ> RemoveShortSegments(List<XYZ> points)
        {
            List<XYZ> result = new List<XYZ>();
            foreach (XYZ point in points)
            {
                if (result.Count == 0 || result.Last().DistanceTo(point) >= MinSegmentLength)
                    result.Add(point);
            }

            return result;
        }

        private bool HasValidSegments(List<XYZ> points)
        {
            for (int i = 1; i < points.Count; i++)
            {
                if (points[i - 1].DistanceTo(points[i]) < MinSegmentLength)
                    return false;
            }

            return true;
        }

        private sealed class ProjectionRange
        {
            public ProjectionRange(double min, double max)
            {
                Min = min;
                Max = max;
            }

            public double Min { get; }

            public double Max { get; }
        }
    }
}
