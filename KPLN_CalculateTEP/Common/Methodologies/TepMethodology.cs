using Autodesk.Revit.DB;

namespace KPLN_CalculateTEP.Common.Methodologies
{
    // Existing differences only. Shared calculations remain in Engine partials.
    internal abstract class TepMethodology
    {
        private static readonly TepMethodology Old = new OldMethodologyCalculation();
        private static readonly TepMethodology New = new NewMethodologyCalculation();

        internal static TepMethodology For(string method) => method == "new" ? New : Old;
        internal abstract string GnsWallBoundary { get; }
        internal abstract bool UsesCoreBoundary { get; }
        internal abstract bool ExcludesGnsRole(string role);
        internal bool CountsGnsOpeningOnOneFloor(Solid shape) => CountsGnsOpeningAreaOnOneFloor(shape.Volume * .09290304);
        internal abstract bool CountsGnsOpeningAreaOnOneFloor(double areaSquareMetres);
        internal abstract bool CountsGnsShaftOnOneFloor { get; }
    }
}
