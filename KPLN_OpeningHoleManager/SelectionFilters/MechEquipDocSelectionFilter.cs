using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_OpeningHoleManager.Forms.SelectionFilters
{
    internal sealed class MechEquipDocSelectionFilter : ISelectionFilter
    {
        public MechEquipDocSelectionFilter()
        {
        }

        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance fi)
                return fi.Symbol.FamilyName.StartsWith("199_Отвер") 
                    || fi.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие");

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
