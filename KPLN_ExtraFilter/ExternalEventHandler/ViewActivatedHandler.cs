using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExternalEventHandler
{
    public sealed class ViewActivatedHandler : IExternalEventHandler
    {
        public SelectionByModelVM CurrentSelByModelVM { get; set; }

        public SetParamsByFrameVM CurrentSetParamsByFrameVM { get; set; }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return;

            Document doc = uidoc.Document;

            // Пользовательский элемент
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            
            // Основные данные для поиска
            IEnumerable<Element> userSelElems = selectedIds.Select(id => doc.GetElement(id));

            if (CurrentSelByModelVM != null)
            {
                CurrentSelByModelVM.CurrentSelectionByModelM.Doc = doc;
                CurrentSelByModelVM.CurrentSelectionByModelM.DocActiveView = doc.ActiveView;
            }

            if (CurrentSetParamsByFrameVM != null)
            {
                CurrentSetParamsByFrameVM.CurrentSetParamsByFrameM.Doc = doc;
                CurrentSetParamsByFrameVM.CurrentSetParamsByFrameM.DocActiveView = doc.ActiveView;
            }
        }

        public string GetName() => "ViewActivatedHandler";
    }
}
