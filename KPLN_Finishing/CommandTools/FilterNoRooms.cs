using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;

namespace KPLN_Finishing.CommandTools
{
    class FilterNoRooms : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
                Parameter roomIdParam = elem.LookupParameter("О_Id помещения");

            if (roomIdParam == null) return false;

            if ((roomIdParam.AsString() == "" || roomIdParam.AsString() == null) && ((elem.Category.Id.IntegerValue == -2000011 || elem.Category.Id.IntegerValue == -2000038 || elem.Category.Id.IntegerValue == -2000032)))
            {
                if (elem.Document.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().ToLower() != "отделка")
                    return false;
                
                return true;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
