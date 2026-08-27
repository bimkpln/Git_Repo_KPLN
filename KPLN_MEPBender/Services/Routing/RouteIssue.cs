using Autodesk.Revit.DB;
using KPLN_MEPBender.Common;

namespace KPLN_MEPBender.Services.Routing
{
    public sealed class RouteIssue
    {
        public RouteIssue(ElementId elementId, string elementType, string stage, string message)
        {
            ElementId = elementId;
            ElementType = elementType;
            Stage = stage;
            Message = message;
        }

        public ElementId ElementId { get; }

        public string ElementType { get; }

        public string Stage { get; }

        public string Message { get; }

        public override string ToString()
        {
            string idValue = ElementId == null || ElementId == Autodesk.Revit.DB.ElementId.InvalidElementId
                ? "-"
                : ElementId.GetStableIntegerValue().ToString();

            return $"id {idValue} ({ElementType}) [{Stage}]: {Message}";
        }
    }
}
