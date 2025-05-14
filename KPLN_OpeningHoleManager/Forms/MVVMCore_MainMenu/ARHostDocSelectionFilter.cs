﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using KPLN_OpeningHoleManager.Services;

namespace KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu
{
    internal sealed class ARHostDocSelectionFilter : ISelectionFilter
    {
        public ARHostDocSelectionFilter()
        {
        }

        public bool AllowElement(Element elem)
        {
            if (elem is Wall wall)
            {
                string typeParamData = wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                bool isMatchTypeNameFilter = ARKRElemsWorker.ElemCatLogicalOrFilter.PassesFilter(elem);
                bool isMatchExtraFilter = ARKRElemsWorker.ElemExtraFilterFunc(elem);

                return isMatchTypeNameFilter && isMatchExtraFilter;
            }

            return false;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
