using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using KPLN_MEPBender.Services.Geometry;

namespace KPLN_MEPBender.Common
{
    public enum MepBenderSelectionMode
    {
        Obstacle,
        Route
    }

    public enum MepBenderObstacleSelectionSource
    {
        Model,
        Link
    }

    public sealed class MepBenderSelectionFilter : ISelectionFilter
    {
        private readonly Document _hostDoc;
        private readonly MepBenderSelectionMode _mode;
        private readonly MepBenderObstacleSelectionSource _obstacleSource;

        public MepBenderSelectionFilter(
            Document hostDoc,
            MepBenderSelectionMode mode,
            MepBenderObstacleSelectionSource obstacleSource = MepBenderObstacleSelectionSource.Model)
        {
            _hostDoc = hostDoc;
            _mode = mode;
            _obstacleSource = obstacleSource;
        }

        public bool AllowElement(Element elem)
        {
            if (elem == null)
                return false;

            if (_mode == MepBenderSelectionMode.Obstacle)
            {
                if (_obstacleSource == MepBenderObstacleSelectionSource.Link)
                    return elem is RevitLinkInstance;

                return !(elem is RevitLinkInstance) && IsObstacleCandidate(elem);
            }

            if (elem.Category == null)
                return false;

            BuiltInCategory category = (BuiltInCategory)elem.Category.Id.GetStableIntegerValue();
            return category == BuiltInCategory.OST_PipeCurves
                   || category == BuiltInCategory.OST_DuctCurves
                   || category == BuiltInCategory.OST_CableTray;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            if (_mode == MepBenderSelectionMode.Obstacle)
                return IsObstacleReference(reference);

            return reference != null && reference.LinkedElementId == ElementId.InvalidElementId;
        }

        private bool IsObstacleReference(Reference reference)
        {
            if (reference == null)
                return false;

            if (reference.LinkedElementId != null && reference.LinkedElementId != ElementId.InvalidElementId)
            {
                if (_obstacleSource != MepBenderObstacleSelectionSource.Link)
                    return false;

                RevitLinkInstance linkInstance = _hostDoc.GetElement(reference) as RevitLinkInstance
                                                    ?? _hostDoc.GetElement(reference.ElementId) as RevitLinkInstance;
                Document linkedDoc = linkInstance?.GetLinkDocument();
                Element linkedElement = linkedDoc?.GetElement(reference.LinkedElementId);
                return IsObstacleCandidate(linkedElement);
            }

            if (_obstacleSource != MepBenderObstacleSelectionSource.Model)
                return false;

            Element hostElement = _hostDoc.GetElement(reference.ElementId);
            return IsObstacleCandidate(hostElement);
        }

        private bool IsObstacleCandidate(Element element)
        {
            return element != null
                   && !(element is ElementType)
                   && !element.ViewSpecific
                   && element.Category != null;
        }
    }
}
