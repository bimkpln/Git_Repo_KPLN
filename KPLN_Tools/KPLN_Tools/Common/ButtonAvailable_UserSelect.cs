using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    /// <summary>
    /// Общий класс для определения "нажимаемости" кнопки в зависимости от наличия выделенных 
    /// </summary>
    internal class ButtonAvailable_UserSelect : IExternalCommandAvailability
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
                }
            }
        }

        public static UIDocument UIDoc
        {
            get => _uidoc;
            private set
            {
                if (value != _uidoc)
                {
                    _uidoc = value;
                }
            }
        }


        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            
            UIAppData = applicationData;
            UIDoc = UIAppData?.ActiveUIDocument;

            if (UIDoc != null)
            {
                Selection selection = UIDoc.Selection;
                if (selection != null)
                {
#if Revit2020 || Debug2020
                    // Для 2020 - нет возможности получить элементы из связи. Это исправили в более поздних версиях
                    ICollection<ElementId> selIds = selection.GetElementIds();
                    if (selIds.Count == 0)
                        return false;
#else
                    IList<Reference> selRefers = selection.GetReferences();
                    if (selRefers.Count == 0)
                        return false;
#endif

                    // У элементов линков CategorySet будет пустым
                    if (selectedCategories.Size == 0)
                        return true;

                    foreach (Category c in selectedCategories)
                    {
#if Revit2020 || Debug2020 || Revit2023 || Debug2023
                        if (c.Id.IntegerValue != (int)BuiltInCategory.OST_Views && c.Id.IntegerValue != (int)BuiltInCategory.OST_Sheets)
#else
                        if (c.BuiltInCategory != BuiltInCategory.OST_Views && c.BuiltInCategory != BuiltInCategory.OST_Sheets)
#endif
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
