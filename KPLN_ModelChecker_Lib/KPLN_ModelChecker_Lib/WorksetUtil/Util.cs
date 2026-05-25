using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.WorksetUtil
{
    public static class Util
    {
        public static Workset[] GetDocWorksets(Document doc)
        {
            if (!doc.IsWorkshared)
                return null;

            return new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .ToArray(); ;
        }

        public static int CountWSElems(Document doc, Workset ws)
        {
            ElementWorksetFilter wfilter = new ElementWorksetFilter(ws.Id);
            FilteredElementCollector col = new FilteredElementCollector(doc).WherePasses(wfilter);

            return col.GetElementCount();
        }

        /// <summary>
        /// Метод для поиска и вывода пользователю пустых рабочих наборов
        /// </summary>
        public static Workset[] GetEmptyWorksets(Document doc)
        {
            List<Workset> result = new List<Workset>();

            foreach (Workset w in GetDocWorksets(doc))
            {
                if (CountWSElems(doc, w) == 0)
                    result.Add(w);
            }

            return result.ToArray();
        }
    }
}
