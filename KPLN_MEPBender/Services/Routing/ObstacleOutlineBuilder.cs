using Autodesk.Revit.DB;
using KPLN_MEPBender.Services.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class ObstacleOutlineBuilder
    {
        private const double SignificantZOffsetRatio = 0.01;
        private const double PlanarFaceMinZNormal = 0.05;
        private const double FaceIntersectionTolerance = 0.001;

        public Outline Build(MepBendRequest request, double expandBy)
        {
            Outline result = null;

            foreach (LinkedElementReference obstacleReference in request.ObstacleReferences)
            {
                Element element = GetSourceElement(request.Doc, obstacleReference);
                BoundingBoxXYZ box = element?.get_BoundingBox(null);
                Outline outline = LinkTransformHelper.ToHostOutline(box, obstacleReference.TransformToHost);
                if (outline == null)
                    continue;

                AddOutline(ref result, outline.MinimumPoint, outline.MaximumPoint);
            }

            if (result == null)
                return null;

            result.AddPoint(result.MinimumPoint - new XYZ(expandBy, expandBy, expandBy));
            result.AddPoint(result.MaximumPoint + new XYZ(expandBy, expandBy, expandBy));

            return result;
        }

        public Outline BuildForRoute(MepBendRequest request, MEPCurve route, double intersectionTolerance, double expandBy)
        {
            BoundingBoxXYZ routeBox = route?.get_BoundingBox(null);
            Outline routeOutline = LinkTransformHelper.ToHostOutline(routeBox, Transform.Identity);
            if (routeOutline == null)
                return null;

            Outline result = null;

            foreach (LinkedElementReference obstacleReference in request.ObstacleReferences)
            {
                Element element = GetSourceElement(request.Doc, obstacleReference);
                BoundingBoxXYZ box = element?.get_BoundingBox(null);
                Outline obstacleOutline = LinkTransformHelper.ToHostOutline(box, obstacleReference.TransformToHost);
                if (obstacleOutline == null)
                    continue;

                if (!routeOutline.Intersects(obstacleOutline, intersectionTolerance))
                    continue;

                obstacleOutline = RefineSlopedObstacleOutline(route, obstacleReference, element, obstacleOutline);
                AddOutline(ref result, obstacleOutline.MinimumPoint, obstacleOutline.MaximumPoint);
            }

            if (result == null)
                return null;

            result.AddPoint(result.MinimumPoint - new XYZ(expandBy, expandBy, expandBy));
            result.AddPoint(result.MaximumPoint + new XYZ(expandBy, expandBy, expandBy));

            return result;
        }

        private Outline RefineSlopedObstacleOutline(MEPCurve route, LinkedElementReference obstacleReference, Element element, Outline obstacleOutline)
        {
            if (!HasSignificantZOffset(obstacleOutline))
                return obstacleOutline;

            LocationCurve locationCurve = route?.Location as LocationCurve;
            Line routeLine = locationCurve?.Curve as Line;
            if (routeLine == null)
                return obstacleOutline;

            if (!TryGetLocalSolidZRange(element, obstacleReference, routeLine, obstacleOutline, out double minZ, out double maxZ))
                return obstacleOutline;

            XYZ min = obstacleOutline.MinimumPoint;
            XYZ max = obstacleOutline.MaximumPoint;
            return new Outline(
                new XYZ(min.X, min.Y, Math.Min(minZ, maxZ)),
                new XYZ(max.X, max.Y, Math.Max(minZ, maxZ)));
        }

        private bool HasSignificantZOffset(Outline outline)
        {
            XYZ min = outline.MinimumPoint;
            XYZ max = outline.MaximumPoint;
            double zDelta = Math.Abs(max.Z - min.Z);
            double horizontalLength = Math.Sqrt(Math.Pow(max.X - min.X, 2) + Math.Pow(max.Y - min.Y, 2));

            if (horizontalLength < 0.001)
                return false;

            return zDelta / horizontalLength > SignificantZOffsetRatio;
        }

        private bool TryGetLocalSolidZRange(
            Element element,
            LinkedElementReference obstacleReference,
            Line routeLine,
            Outline obstacleOutline,
            out double minZ,
            out double maxZ)
        {
            minZ = obstacleOutline.MinimumPoint.Z;
            maxZ = obstacleOutline.MaximumPoint.Z;

            List<Solid> solids = GetHostSolids(element, obstacleReference).ToList();
            if (solids.Count == 0)
                return false;

            foreach (XYZ samplePoint in GetRouteSamplePoints(routeLine, obstacleOutline))
            {
                List<double> lowerZValues = new List<double>();
                List<double> upperZValues = new List<double>();

                foreach (Solid solid in solids)
                    AddIntersectedFaceZValues(solid, samplePoint, lowerZValues, upperZValues);

                bool hasLower = lowerZValues.Count > 0;
                bool hasUpper = upperZValues.Count > 0;
                if (!hasLower && !hasUpper)
                    continue;

                if (hasLower)
                    minZ = lowerZValues.Min();
                if (hasUpper)
                    maxZ = upperZValues.Max();

                return true;
            }

            return false;
        }

        private IEnumerable<XYZ> GetRouteSamplePoints(Line routeLine, Outline obstacleOutline)
        {
            XYZ start = routeLine.GetEndPoint(0);
            XYZ end = routeLine.GetEndPoint(1);
            XYZ direction = end - start;
            double length = direction.GetLength();
            if (length < 0.001)
            {
                yield return start;
                yield break;
            }

            direction = direction.Normalize();
            ProjectionRange range = GetProjectionRange(start, direction, obstacleOutline);
            double min = Clamp(range.Min, 0, length);
            double max = Clamp(range.Max, 0, length);
            if (max < min)
            {
                double buffer = min;
                min = max;
                max = buffer;
            }

            double[] factors = { 0.5, 0.25, 0.75, 0, 1 };
            foreach (double factor in factors)
                yield return start + direction * (min + (max - min) * factor);
        }

        private ProjectionRange GetProjectionRange(XYZ start, XYZ direction, Outline outline)
        {
            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (XYZ point in GetOutlineCorners(outline))
            {
                double projection = (point - start).DotProduct(direction);
                min = Math.Min(min, projection);
                max = Math.Max(max, projection);
            }

            return new ProjectionRange(min, max);
        }

        private void AddIntersectedFaceZValues(Solid solid, XYZ samplePoint, List<double> lowerZValues, List<double> upperZValues)
        {
            foreach (Face face in solid.Faces)
            {
                PlanarFace planarFace = face as PlanarFace;
                if (planarFace == null)
                    continue;

                XYZ normal = planarFace.FaceNormal;
                if (normal == null || Math.Abs(normal.Z) < PlanarFaceMinZNormal)
                    continue;

                if (!TryGetVerticalIntersectionZ(planarFace, samplePoint, out double z))
                    continue;

                if (normal.Z < 0)
                    lowerZValues.Add(z);
                else
                    upperZValues.Add(z);
            }
        }

        private bool TryGetVerticalIntersectionZ(PlanarFace face, XYZ samplePoint, out double z)
        {
            z = 0;
            XYZ normal = face.FaceNormal;
            if (normal == null || Math.Abs(normal.Z) < PlanarFaceMinZNormal)
                return false;

            double planeOffset = normal.DotProduct(face.Origin);
            z = (planeOffset - normal.X * samplePoint.X - normal.Y * samplePoint.Y) / normal.Z;
            XYZ pointOnPlane = new XYZ(samplePoint.X, samplePoint.Y, z);
            IntersectionResult projection = face.Project(pointOnPlane);
            if (projection == null || projection.XYZPoint == null)
                return false;

            return projection.XYZPoint.DistanceTo(pointOnPlane) <= FaceIntersectionTolerance;
        }

        private IEnumerable<Solid> GetHostSolids(Element element, LinkedElementReference obstacleReference)
        {
            if (element == null)
                yield break;

            Options options = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geometryElement = element.get_Geometry(options);
            if (geometryElement == null)
                yield break;

            foreach (Solid solid in GetSolids(geometryElement))
            {
                Solid hostSolid = LinkTransformHelper.ToHostSolid(obstacleReference, solid);
                if (IsValidSolid(hostSolid))
                    yield return hostSolid;
            }
        }

        private IEnumerable<Solid> GetSolids(GeometryElement geometryElement)
        {
            foreach (GeometryObject geometryObject in geometryElement)
            {
                Solid solid = geometryObject as Solid;
                if (IsValidSolid(solid))
                {
                    yield return solid;
                    continue;
                }

                GeometryInstance geometryInstance = geometryObject as GeometryInstance;
                GeometryElement instanceGeometry = geometryInstance?.GetInstanceGeometry();
                if (instanceGeometry == null)
                    continue;

                foreach (Solid nestedSolid in GetSolids(instanceGeometry))
                {
                    if (IsValidSolid(nestedSolid))
                        yield return nestedSolid;
                }
            }
        }

        private bool IsValidSolid(Solid solid)
        {
            return solid != null && solid.Faces.Size > 0 && solid.Edges.Size > 0;
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

        private double Clamp(double value, double min, double max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private Element GetSourceElement(Document hostDoc, LinkedElementReference obstacleReference)
        {
            if (!obstacleReference.IsLinked)
                return hostDoc.GetElement(obstacleReference.HostElementId);

            RevitLinkInstance linkInstance = hostDoc.GetElement(obstacleReference.HostElementId) as RevitLinkInstance;
            Document linkedDoc = linkInstance?.GetLinkDocument();
            return linkedDoc?.GetElement(obstacleReference.LinkedElementId);
        }

        private void AddOutline(ref Outline outline, XYZ min, XYZ max)
        {
            if (outline == null)
            {
                outline = new Outline(min, max);
                return;
            }

            outline.AddPoint(min);
            outline.AddPoint(max);
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