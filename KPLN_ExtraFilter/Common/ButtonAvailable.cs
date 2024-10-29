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
                if (c.Id.IntegerValue != (int)BuiltInCategory.OST_Views)
                    return true;
            }

            return false;
        }
    }
}
