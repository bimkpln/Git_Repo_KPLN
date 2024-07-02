using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;

namespace KPLN_Tools.Common.SS_System
{
    internal class PickFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            Category category = elem.Category;
            if (category != null)
            {
                int elemCatId = category.Id.IntegerValue;
                if (elemCatId == (int)BuiltInCategory.OST_ElectricalEquipment)
                    return true;

            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }
}
