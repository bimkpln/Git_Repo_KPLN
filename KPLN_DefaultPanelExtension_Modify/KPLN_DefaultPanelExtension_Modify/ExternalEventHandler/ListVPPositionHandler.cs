using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_DefaultPanelExtension_Modify.ExternalEventHandler
{
    public sealed class ListVPPositionHandler : IExternalEventHandler
    {
        public ListVPPositionMainVM HandlerListVPPositionMainVM { get; set; }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return;

            Document doc = uidoc.Document;

            // Пользовательский элемент
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

            // Основные данные для анализа
            IEnumerable<Element> userSelElems = selectedIds
                .Select(id => doc.GetElement(id))
                .Where(el => el is Viewport || el is ScheduleSheetInstance);
            
            if (HandlerListVPPositionMainVM != null)
                HandlerListVPPositionMainVM.SelectedViewElems = userSelElems.ToArray();
        }

        public string GetName() => "ListVPPositionHandler";
    }
}
