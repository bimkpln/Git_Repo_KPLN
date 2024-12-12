using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Looker.Core
{
    /// <summary>
    /// Класс для сравнения элементов по ID (для создания HashSet)
    /// </summary>
    internal sealed class ElementComparerById : IEqualityComparer<Element>
    {
        public bool Equals(Element x, Element y) => x.Id == y.Id;

        public int GetHashCode(Element obj) => obj.Id.GetHashCode();
    }
}
