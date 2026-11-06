using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_OpeningHoleManager.Core
{
    internal sealed class OpeningHistoryChanges
    {
        internal Dictionary<string, ElementId> Elements { get; } = new Dictionary<string, ElementId>();
        internal List<FamilyInstance> Created { get; } = new List<FamilyInstance>();

        internal bool RefreshLevels { get; set; }
        internal bool RefreshVisibility { get; set; }
        internal bool FullRefresh { get; set; }
        internal bool CancelAutoCache { get; set; }

        internal bool HasChanges => Elements.Count > 0 || RefreshLevels || RefreshVisibility || FullRefresh;


        internal void Add(ElementId id)
        {
            Elements[id.ToString()] = id;
        }
    }
}
