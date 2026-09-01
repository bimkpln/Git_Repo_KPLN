using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_TrailingMEP.Common
{
    internal sealed class RouteCurveSelectionFilter : ISelectionFilter
    {
        private readonly string _lineStyleName;

        public RouteCurveSelectionFilter(string lineStyleName)
        {
            _lineStyleName = lineStyleName;
        }

        public bool AllowElement(Element elem)
        {
            if (!(elem is CurveElement curveElement))
                return false;

            if (!(curveElement.GeometryCurve is Line))
                return false;

            GraphicsStyle lineStyle = curveElement.LineStyle as GraphicsStyle;
            if (lineStyle == null)
                return false;

            return lineStyle.Name == _lineStyleName
                || lineStyle.GraphicsStyleCategory?.Name == _lineStyleName;
        }

        public bool AllowReference(Reference reference, XYZ position) => true;
    }
}
