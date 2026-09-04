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
            double offsetIterationStepMm,
            double angleDegrees,
            IReadOnlyCollection<BendDirection> directions,
            bool alignVerticalBendByLowest,
            bool analyzeCollisions)
        {
            Doc = doc;
            ActiveView = activeView;
            RouteElementIds = new List<ElementId>(routeElementIds);
            ObstacleReferences = new List<LinkedElementReference>(obstacleReferences);
            OffsetMm = offsetMm;
            OffsetIterationStepMm = offsetIterationStepMm;
            AngleDegrees = angleDegrees;
            Directions = directions;
            AlignVerticalBendByLowest = alignVerticalBendByLowest;
            AnalyzeCollisions = analyzeCollisions;
        }

        public Document Doc { get; }

        public View ActiveView { get; }

        public IReadOnlyCollection<ElementId> RouteElementIds { get; }

        public IReadOnlyCollection<LinkedElementReference> ObstacleReferences { get; }

        public double OffsetMm { get; }

        public double OffsetIterationStepMm { get; }

        public double AngleDegrees { get; }

        public IReadOnlyCollection<BendDirection> Directions { get; }

        public bool AlignVerticalBendByLowest { get; }

        public bool AnalyzeCollisions { get; }

        public MepBendRequest WithOffset(double offsetMm)
        {
            return new MepBendRequest(
                Doc,
                ActiveView,
                RouteElementIds,
                ObstacleReferences,
                offsetMm,
                OffsetIterationStepMm,
                AngleDegrees,
                Directions,
                AlignVerticalBendByLowest,
                AnalyzeCollisions);
        }
    }
}