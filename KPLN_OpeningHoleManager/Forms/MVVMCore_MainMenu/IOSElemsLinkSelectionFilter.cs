using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using KPLN_OpeningHoleManager.Services;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal sealed class IOSElemsLinkSelectionFilter : ISelectionFilter
    {
        private readonly Document _doc;
        private Document _linkDoc = null;

        public IOSElemsLinkSelectionFilter(Document doc)
        {
            _doc = doc;
        }

        public bool AllowElement(Element elem) => true;

        public bool AllowReference(Reference reference, XYZ position)
        {
            if (_doc.GetElement(reference) is RevitLinkInstance rli)
                _linkDoc = rli.GetLinkDocument();
            else
            {
                _linkDoc = null;
                return false;
            }

            if (_linkDoc.GetElement(reference.LinkedElementId) is Element elem)
            {
                bool isMatchCatFilter = IOSElemsCollectionCreator.ElemCatLogicalOrFilter.PassesFilter(elem);
                bool isMatchExtraFilter = IOSElemsCollectionCreator.ElemExtraFilterFunc(elem);

                return isMatchCatFilter && isMatchExtraFilter;
            }

            return false;
        }
    }
}
