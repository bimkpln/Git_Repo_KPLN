using Autodesk.Revit.DB;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class OpenDocumentItem
    {
        internal Document Document
        { get; set; }

        public string DisplayName
        { get; set; }

        public string LoadTargetName
        { get; set; }
    }
}
