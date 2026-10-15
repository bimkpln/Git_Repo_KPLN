using Autodesk.Revit.DB;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Положение аннотации на печатной плоскости; глубину назначает Revit.</summary>
    internal static class AnnotationPlane
    {
        internal static XYZ KeepDepth(XYZ requested, XYZ actual, XYZ normal)
        {
            return requested + normal * ((actual - requested).DotProduct(normal));
        }

        internal static double Distance(XYZ requested, XYZ actual, XYZ normal)
        {
            return KeepDepth(requested, actual, normal).DistanceTo(actual);
        }
    }
}
