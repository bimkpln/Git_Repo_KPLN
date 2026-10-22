using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Tools_Lib.Services
{
    internal sealed class CreateSpacesResult
    {
        internal int Created { get; set; }
        internal int Existing { get; set; }
        internal int Skipped { get; set; }
        internal List<string> Messages { get; } = new List<string>();

        internal Dictionary<string, List<ElementId>> FailureIdsByMessage { get; } = new Dictionary<string, List<ElementId>>();

        internal string Summary => string.Format(
            "Создано пространств: {0}. Уже существуют: {1}. Пропущено помещений: {2}.",
            Created, Existing, Skipped);
    }
}
