using Autodesk.Revit.DB;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class ElementLocationService
    {
        private static double[] Point(XYZ point) => new[] { point.X * 304.8, point.Y * 304.8, point.Z * 304.8 };

        internal static object Read(Element element)
        {
            var point = element.Location as LocationPoint;
            if (point != null)
            {
                // Группы, помещения и ряд аннотаций имеют точку, но не поддерживают Rotation.
                double? rotation = null;
                string rotationStatus = "available";
                try
                {
                    rotation = point.Rotation;
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    rotationStatus = "not_supported";
                }

                return new { kind = "point", point_mm = Point(point.Point), rotation_radians = rotation, rotation_status = rotationStatus };
            }

            var curve = (element.Location as LocationCurve)?.Curve;
            if (curve != null && curve.IsBound)
                return new { kind = "curve", curve_type = curve.GetType().Name, start_mm = Point(curve.GetEndPoint(0)), end_mm = Point(curve.GetEndPoint(1)), length_mm = curve.Length * 304.8 };

            return null;
        }
    }
}
