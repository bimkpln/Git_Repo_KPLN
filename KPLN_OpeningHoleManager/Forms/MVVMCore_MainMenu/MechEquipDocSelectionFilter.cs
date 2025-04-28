using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal sealed class MechEquipDocSelectionFilter : ISelectionFilter
    {
        public MechEquipDocSelectionFilter()
        {
        }

        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance fi)
                return fi.Symbol.FamilyName.StartsWith("199_") || fi.Symbol.FamilyName.StartsWith("ASML_АР_Отверстие");

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
