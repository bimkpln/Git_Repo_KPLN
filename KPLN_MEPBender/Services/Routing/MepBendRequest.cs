using Autodesk.Revit.DB;
using KPLN_MEPBender.Services.Geometry;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Routing
{
    public sealed class MepBendRequest
    {
        public MepBendRequest(
            Document doc,
            View activeView,
            IEnumerable<ElementId> routeElementIds,
            IEnumerable<LinkedElementReference> obstacleReferences,
            double offsetMm,
            double angleDegrees,
            IReadOnlyCollection<BendDirection> directions,
            bool analyzeCollisions)
        {
            Doc = doc;
            ActiveView = activeView;
            RouteElementIds = new List<ElementId>(routeElementIds);
            ObstacleReferences = new List<LinkedElementReference>(obstacleReferences);
            OffsetMm = offsetMm;
            AngleDegrees = angleDegrees;
            Directions = directions;
            AnalyzeCollisions = analyzeCollisions;
        }

        public Document Doc { get; }

        public View ActiveView { get; }

        public IReadOnlyCollection<ElementId> RouteElementIds { get; }

        public IReadOnlyCollection<LinkedElementReference> ObstacleReferences { get; }

        public double OffsetMm { get; }

        public double AngleDegrees { get; }

        public IReadOnlyCollection<BendDirection> Directions { get; }

        public bool AnalyzeCollisions { get; }
    }
}
