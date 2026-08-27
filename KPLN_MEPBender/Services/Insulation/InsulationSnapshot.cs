using Autodesk.Revit.DB;
using KPLN_MEPBender.Services.Parameters;

namespace KPLN_MEPBender.Services.Insulation
{
    internal enum InsulationSnapshotKind
    {
        PipeInsulation,
        DuctInsulation,
        DuctLining
    }

    internal sealed class InsulationSnapshot
    {
        public InsulationSnapshot(InsulationSnapshotKind kind, ElementId typeId, double thickness, ParameterSnapshot parameters)
        {
            Kind = kind;
            TypeId = typeId;
            Thickness = thickness;
            Parameters = parameters;
        }

        public InsulationSnapshotKind Kind { get; }

        public ElementId TypeId { get; }

        public double Thickness { get; }

        public ParameterSnapshot Parameters { get; }
    }
}
