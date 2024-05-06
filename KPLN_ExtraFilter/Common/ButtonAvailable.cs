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
                if (value != _uiAppData)
                {
                    _uiAppData = value;
                    _uidoc = _uiAppData?.ActiveUIDocument;
                }
            }
        }

        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            UIAppData = applicationData;

            if (_uidoc != null)
            {
                Selection selection = _uidoc.Selection;
                if (selection != null)
                {
                    ICollection<ElementId> selectedIds = selection.GetElementIds();
                    if (selectedIds.Count == 1)
                    {
                        foreach (Category c in selectedCategories)
                        {
                            if (c.Id.IntegerValue != (int)BuiltInCategory.OST_Views)
                                return true;
                        }
                    }
                }
            }

            return false;
        }
    }
}
