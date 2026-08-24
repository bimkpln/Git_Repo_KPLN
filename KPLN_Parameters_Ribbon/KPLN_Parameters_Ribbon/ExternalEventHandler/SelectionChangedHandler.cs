using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Forms.ViewModels;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Parameters_Ribbon.ExternalEventHandler
{
    public sealed class SelectionChangedHandler : IExternalEventHandler
    {
        public SumParametersVM CurrentSumParametersVM { get; set; }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null || CurrentSumParametersVM == null)
                return;

            Document doc = uidoc.Document;
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            IEnumerable<Element> userSelElems = selectedIds
                .Select(id => doc.GetElement(id))
                .Where(e => e != null);

            CurrentSumParametersVM.SetUserSelection(doc, userSelElems);
        }

        public string GetName() => "SelectionChangedHandler";
    }
}
