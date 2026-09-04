using Autodesk.Revit.DB;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class MepBenderFailure
    {
        public MepBenderFailure(FailureDefinitionId failureDefinitionId, string description)
        {
            FailureDefinitionId = failureDefinitionId;
            Description = description ?? string.Empty;
        }

        public FailureDefinitionId FailureDefinitionId { get; }

        public string Description { get; }
    }
}