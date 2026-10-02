using Autodesk.Revit.DB;

namespace KPLN_CalculateTEP.Common.Methodologies
{
    /// <summary>Существующие правила старой методики, перенесённые без изменения.</summary>
    internal sealed class OldMethodologyCalculation : TepMethodology
    {
        internal override string GnsWallBoundary => "exterior";
        internal override bool UsesCoreBoundary => false;
        internal override bool ExcludesGnsRole(string role) => false;
        internal override bool CountsGnsOpeningOnOneFloor(Solid shape) => shape.Volume * .09290304 > 36;
        internal override bool CountsGnsShaftOnOneFloor => false;
    }
}
