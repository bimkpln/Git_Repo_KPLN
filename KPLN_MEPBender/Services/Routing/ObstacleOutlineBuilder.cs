using Autodesk.Revit.DB;
using KPLN_MEPBender.Services.Geometry;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class ObstacleOutlineBuilder
    {
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

                AddOutline(ref result, obstacleOutline.MinimumPoint, obstacleOutline.MaximumPoint);
            }

            if (result == null)
                return null;

            result.AddPoint(result.MinimumPoint - new XYZ(expandBy, expandBy, expandBy));
            result.AddPoint(result.MaximumPoint + new XYZ(expandBy, expandBy, expandBy));

            return result;
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
    }
}
