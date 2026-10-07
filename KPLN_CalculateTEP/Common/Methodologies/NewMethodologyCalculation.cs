using Autodesk.Revit.DB;

namespace KPLN_CalculateTEP.Common.Methodologies
{
    /// <summary>Существующие правила новой методики, перенесённые без изменения.</summary>
    internal sealed class NewMethodologyCalculation : TepMethodology
    {
        internal override string GnsWallBoundary => "core";
        internal override bool UsesCoreBoundary => true;
        internal override bool ExcludesGnsRole(string role) => role == "roof-used" || role == "terrace";
        internal override bool CountsGnsOpeningAreaOnOneFloor(double areaSquareMetres) => true;
        internal override bool CountsGnsShaftOnOneFloor => true;
    }
}
