using Autodesk.Revit.DB;

namespace KPLN_MEPBender.Services.Routing
{
    internal sealed class ExternalConnectorInfo
    {
        public ExternalConnectorInfo(ElementId ownerId, XYZ sourceEndpoint)
        {
            OwnerId = ownerId;
            SourceEndpoint = sourceEndpoint;
        }

        public ElementId OwnerId { get; }

        public XYZ SourceEndpoint { get; }
    }
}
