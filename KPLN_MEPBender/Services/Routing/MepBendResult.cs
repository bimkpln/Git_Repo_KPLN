using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Routing
{
    public sealed class MepBendResult
    {
        public MepBendResult()
        {
            CreatedElementIds = new List<ElementId>();
            ProcessedRouteIds = new List<ElementId>();
            Issues = new List<RouteIssue>();
        }

        public List<ElementId> CreatedElementIds { get; }

        public List<ElementId> ProcessedRouteIds { get; }

        public List<RouteIssue> Issues { get; }

        public int CapturedParameterSnapshotsCount { get; set; }

        public int SkippedRouteCount { get; set; }

        public int FailedRouteCount { get; set; }

        public int FittingFailureCount { get; set; }

        public int ReconnectFailureCount { get; set; }

        public bool GeometryWasChanged { get; set; }

        public bool HasInvalidFittingFamilyFailure { get; set; }

        public bool HasInsufficientSpaceFailure { get; set; }

        public string Message { get; set; }

        public void AddIssue(ElementId elementId, string elementType, string stage, string message)
        {
            Issues.Add(new RouteIssue(elementId, elementType, stage, message));
        }
    }
}
