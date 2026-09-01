using Autodesk.Revit.DB;

namespace KPLN_TrailingMEP.Common
{
    public enum MepRouteKind
    {
        Pipe,
        Duct,
        CableTray
    }

    /// <summary>
    /// Данные исходного линейного MEP-элемента, нужные для создания продолжения.
    /// </summary>
    public sealed class MepCurveData
    {
        public ElementId SourceId { get; set; }

        public MepRouteKind Kind { get; set; }

        public ElementId TypeId { get; set; }

        public ElementId SystemTypeId { get; set; }

        public ElementId LevelId { get; set; }

        public XYZ ExtensionStart { get; set; }

        public XYZ OppositeEnd { get; set; }

        public XYZ SourceDirection => (ExtensionStart - OppositeEnd).Normalize();
    }
}
