using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using KPLN_Finishing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Finishing.CommandTools
{
    class Filter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            try
            {
                if (elem.Category.Id.IntegerValue == -2000160 || (elem.LookupParameter(Names.parameter_Room_Id).AsString() != "" && (elem.Category.Id.IntegerValue == -2000011 || elem.Category.Id.IntegerValue == -2000038 || elem.Category.Id.IntegerValue == -2000032)))
                { return true; }
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
