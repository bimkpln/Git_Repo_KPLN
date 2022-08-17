using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;

namespace KPLN_Finishing.CommandTools
{
    class FilterNoRooms : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            try
            {
                if ((elem.LookupParameter("О_Id помещения").AsString() == "" || elem.LookupParameter("О_Id помещения").AsString() == null) && ((elem.Category.Id.IntegerValue == -2000011 || elem.Category.Id.IntegerValue == -2000038 || elem.Category.Id.IntegerValue == -2000032)))
                {
                    if (elem.Document.GetElement(elem.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().ToLower() != "отделка")
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
