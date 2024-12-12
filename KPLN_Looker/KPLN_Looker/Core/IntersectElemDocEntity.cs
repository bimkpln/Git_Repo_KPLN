using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_Looker.Core
{
    /// <summary>
    /// Сущность для запаковки данных по элементам, которые нужно проанализировать
    /// </summary>
    internal sealed class IntersectElemDocEntity
    {
        public IntersectElemDocEntity(Document doc, HashSet<Element> currenDocPotentialIntersectElemColl)
        {
            CurrentDoc = doc;
            CurrenDocPotentialIntersectElemColl = currenDocPotentialIntersectElemColl;
        }

        public Document CurrentDoc { get; }

        public HashSet<Element> CurrenDocPotentialIntersectElemColl { get; }
    }
}