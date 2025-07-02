using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_OpeningHoleManager.Forms.SelectionFilters
{
    internal sealed class MechEquipLinkSelectionFilter : ISelectionFilter
    {
        private readonly Document _doc;
        private Document _linkDoc = null;

        public MechEquipLinkSelectionFilter(Document doc)
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

            if (_linkDoc.GetElement(reference.LinkedElementId) is FamilyInstance fi)
                return 
                    fi.Category != null
                    && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_MechanicalEquipment
                    && fi.Symbol.FamilyName.StartsWith("501_ЗИ_Отвер");

            return false;
        }
    }
}
