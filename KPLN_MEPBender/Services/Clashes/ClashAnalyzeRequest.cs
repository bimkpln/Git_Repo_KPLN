using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Clashes
{
    public sealed class ClashAnalyzeRequest
    {
        public ClashAnalyzeRequest(Document doc, IEnumerable<ElementId> createdElementIds, bool isEnabled)
        {
            Doc = doc;
            CreatedElementIds = new List<ElementId>(createdElementIds);
            IsEnabled = isEnabled;
        }

        public Document Doc { get; }

        public IReadOnlyCollection<ElementId> CreatedElementIds { get; }

        public bool IsEnabled { get; }
    }
}
