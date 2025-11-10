using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class SelectionChangedHandler : IExternalEventHandler
    {
        public SelectionByClickVM ViewModel { get; set; }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Пользовательский элемент
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            
            // Основные данные для поиска
            IEnumerable<Element> userSelElems = selectedIds.Select(id => doc.GetElement(id));
            if (ViewModel != null)
            {
                ViewModel.CurrentSelectionByClickM.UserSelDoc = doc;
                ViewModel.CurrentSelectionByClickM.UserSelElems = userSelElems;
            }
        }

        public string GetName() => "SelectionChangedHandler";
    }
}
