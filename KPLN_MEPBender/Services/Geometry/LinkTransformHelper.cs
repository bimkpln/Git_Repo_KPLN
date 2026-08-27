using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Geometry
{
    public static class LinkTransformHelper
    {
        public static LinkedElementReference CreateReference(Document hostDoc, Reference reference)
        {
            if (hostDoc == null || reference == null)
                return null;

            if (reference.LinkedElementId != null && reference.LinkedElementId != ElementId.InvalidElementId)
            {
                RevitLinkInstance linkInstance = hostDoc.GetElement(reference) as RevitLinkInstance
                                                    ?? hostDoc.GetElement(reference.ElementId) as RevitLinkInstance;
                Document linkedDoc = linkInstance?.GetLinkDocument();
                Element linkedElement = linkedDoc?.GetElement(reference.LinkedElementId);

                if (!Has3DGeometry(linkedElement))
                    return null;

                return new LinkedElementReference(
                    reference.ElementId,
                    reference.LinkedElementId,
                    linkInstance?.GetTotalTransform() ?? Transform.Identity,
                    linkedDoc.Title);
            }

            Element hostElement = hostDoc.GetElement(reference.ElementId);
            if (!Has3DGeometry(hostElement))
                return null;

            return new LinkedElementReference(
                reference.ElementId,
                ElementId.InvalidElementId,
                Transform.Identity,
                hostDoc.Title);
        }

        public static XYZ ToHostPoint(LinkedElementReference elementReference, XYZ sourcePoint)
        {
            if (elementReference == null || sourcePoint == null)
                return sourcePoint;

            return elementReference.TransformToHost.OfPoint(sourcePoint);
        }

        public static Solid ToHostSolid(LinkedElementReference elementReference, Solid sourceSolid)
        {
            if (elementReference == null || sourceSolid == null || !elementReference.IsLinked)
                return sourceSolid;

            return SolidUtils.CreateTransformed(sourceSolid, elementReference.TransformToHost);
        }

        public static Outline ToHostOutline(BoundingBoxXYZ sourceBox, Transform transformToHost)
        {
            if (sourceBox == null)
                return null;

            Transform transform = transformToHost ?? Transform.Identity;
            List<XYZ> points = new List<XYZ>
            {
                transform.OfPoint(new XYZ(sourceBox.Min.X, sourceBox.Min.Y, sourceBox.Min.Z)),
                transform.OfPoint(new XYZ(sourceBox.Min.X, sourceBox.Min.Y, sourceBox.Max.Z)),
                transform.OfPoint(new XYZ(sourceBox.Min.X, sourceBox.Max.Y, sourceBox.Min.Z)),
                transform.OfPoint(new XYZ(sourceBox.Min.X, sourceBox.Max.Y, sourceBox.Max.Z)),
                transform.OfPoint(new XYZ(sourceBox.Max.X, sourceBox.Min.Y, sourceBox.Min.Z)),
                transform.OfPoint(new XYZ(sourceBox.Max.X, sourceBox.Min.Y, sourceBox.Max.Z)),
                transform.OfPoint(new XYZ(sourceBox.Max.X, sourceBox.Max.Y, sourceBox.Min.Z)),
                transform.OfPoint(new XYZ(sourceBox.Max.X, sourceBox.Max.Y, sourceBox.Max.Z))
            };

            XYZ min = points[0];
            XYZ max = points[0];

            foreach (XYZ point in points)
            {
                min = new XYZ(
                    Math.Min(min.X, point.X),
                    Math.Min(min.Y, point.Y),
                    Math.Min(min.Z, point.Z));

                max = new XYZ(
                    Math.Max(max.X, point.X),
                    Math.Max(max.Y, point.Y),
                    Math.Max(max.Z, point.Z));
            }

            return new Outline(min, max);
        }

        public static bool Has3DGeometry(Element element)
        {
            if (element == null || element is ElementType || element.ViewSpecific)
                return false;

            Options options = new Options
            {
                ComputeReferences = false,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            GeometryElement geometryElement = element.get_Geometry(options);
            if (geometryElement == null)
                return false;

            foreach (GeometryObject geometryObject in geometryElement)
            {
                if (Has3DGeometry(geometryObject))
                    return true;
            }

            return false;
        }

        private static bool Has3DGeometry(GeometryObject geometryObject)
        {
            Solid solid = geometryObject as Solid;
            if (solid != null && solid.Faces.Size > 0)
                return true;

            Mesh mesh = geometryObject as Mesh;
            if (mesh != null && mesh.NumTriangles > 0)
                return true;

            GeometryInstance geometryInstance = geometryObject as GeometryInstance;
            GeometryElement instanceGeometry = geometryInstance?.GetInstanceGeometry();
            if (instanceGeometry == null)
                return false;

            foreach (GeometryObject nestedObject in instanceGeometry)
            {
                if (Has3DGeometry(nestedObject))
                    return true;
            }

            return false;
        }
    }
}
