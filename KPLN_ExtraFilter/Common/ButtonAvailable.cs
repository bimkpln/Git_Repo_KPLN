using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

namespace KPLN_ExtraFilter.Common
{
    internal class ButtonAvailable : IExternalCommandAvailability
    {
        private static UIApplication _uiAppData;
        private static UIDocument _uidoc;

        public static UIApplication UIAppData
        {
            get => _uiAppData;
            private set
            {
                if (value == _uiAppData) 
                    return;
                
                _uiAppData = value;
                _uidoc = _uiAppData?.ActiveUIDocument;
            }
        }

        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            UIAppData = applicationData;

            Selection selection = _uidoc?.Selection;
            if (selection == null) 
                return false;
                
            ICollection<ElementId> selectedIds = selection.GetElementIds();
            if (selectedIds.Count != 1) 
                return false;
            
            foreach (Category c in selectedCategories)
            {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                if (c.Id.IntegerValue != (int)BuiltInCategory.OST_Views)
#else
                if (c.BuiltInCategory != BuiltInCategory.OST_Views)
#endif
                return true;
            }

            return false;
        }
    }
}
