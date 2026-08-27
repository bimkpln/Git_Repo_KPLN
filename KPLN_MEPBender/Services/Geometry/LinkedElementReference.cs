using Autodesk.Revit.DB;

namespace KPLN_MEPBender.Services.Geometry
{
    public sealed class LinkedElementReference
    {
        public LinkedElementReference(
            ElementId hostElementId,
            ElementId linkedElementId,
            Transform transformToHost,
            string sourceDocumentTitle)
        {
            HostElementId = hostElementId;
            LinkedElementId = linkedElementId;
            TransformToHost = transformToHost ?? Transform.Identity;
            SourceDocumentTitle = sourceDocumentTitle;
        }

        public ElementId HostElementId { get; }

        public ElementId LinkedElementId { get; }

        public Transform TransformToHost { get; }

        public string SourceDocumentTitle { get; }

        public bool IsLinked => LinkedElementId != ElementId.InvalidElementId;

        public ElementId SourceElementId => IsLinked ? LinkedElementId : HostElementId;
    }
}
