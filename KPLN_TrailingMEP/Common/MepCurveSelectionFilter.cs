using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace KPLN_TrailingMEP.Common
{
    internal sealed class MepCurveSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem?.Category == null)
                return false;

            BuiltInCategory category = GetBuiltInCategory(elem.Category.Id);
            return category == BuiltInCategory.OST_PipeCurves
                || category == BuiltInCategory.OST_DuctCurves
                || category == BuiltInCategory.OST_CableTray;
        }

        public bool AllowReference(Reference reference, XYZ position) => true;

        private static BuiltInCategory GetBuiltInCategory(ElementId categoryId)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return (BuiltInCategory)categoryId.IntegerValue;
#else
            return (BuiltInCategory)categoryId.Value;
#endif
        }
    }
}
