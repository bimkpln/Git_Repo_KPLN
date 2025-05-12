using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal sealed class ARHostDocSelectionFilter : ISelectionFilter
    {
        public ARHostDocSelectionFilter()
        {
        }

        public bool AllowElement(Element elem) =>
            (elem is Wall);

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
