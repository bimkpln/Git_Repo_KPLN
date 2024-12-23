using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_IOSClasher.Core
{
    /// <summary>
    /// Класс для сравнения элементов по ID (для создания HashSet)
    /// </summary>
    internal sealed class ElementComparerById : IEqualityComparer<Element>
    {
        public bool Equals(Element x, Element y)
        {
            if (x == null || y == null)
                return false;

            return x.Id == y.Id;
        }

        public int GetHashCode(Element obj) => obj.Id.GetHashCode();
    }
}
