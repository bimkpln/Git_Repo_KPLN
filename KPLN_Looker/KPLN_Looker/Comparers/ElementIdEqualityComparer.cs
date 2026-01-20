using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Looker.Comparers
{
    internal sealed class ElementIdEqualityComparer : IEqualityComparer<ElementId>
    {
        public static readonly ElementIdEqualityComparer Instance = new ElementIdEqualityComparer();
        private ElementIdEqualityComparer() { }

        public bool Equals(ElementId x, ElementId y) => x == y;
        
        public int GetHashCode(ElementId obj) => obj?.GetHashCode() ?? 0;
    }
}
