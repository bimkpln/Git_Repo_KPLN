using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.Forms.Entities.SearchById
{
    public sealed class SearchByIdDocEntity
    {
        /// <summary>
        /// Конструктор для текущей модели
        /// </summary>
        public SearchByIdDocEntity(Document doc, Element[] elems)
        {
            SDE_Doc = doc;
            SDE_DocElems = elems;
            
            string fullDocName = doc.IsWorkshared && !doc.IsDetached
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;
            SDE_ModelName = fullDocName.Split('\\')
                .FirstOrDefault(str => str.Contains(".rvt"))
                .TrimEnd(".rvt".ToCharArray());
        }

        /// <summary>
        /// Конструктор для линка
        /// </summary>
        public SearchByIdDocEntity(Document doc, Element[] elems, RevitLinkInstance rli) : this(doc, elems)
        {
            SDE_RLI = rli;
        }

        public Document SDE_Doc { get; private set; }

        public RevitLinkInstance SDE_RLI { get; private set; }

        public string SDE_ModelName { get; private set; }

        public Element[] SDE_DocElems { get; private set; }

        public List<Element> GetElementsFromModelById(string strIds)
        {
            List<Element> result = new List<Element>();
            
            string[] strIdsArr;
            if (strIds.Contains(","))
                strIdsArr = strIds.Split(',');
            else
                strIdsArr = new string[] { strIds };

            foreach (string strId in strIdsArr) 
            {
                if (!int.TryParse(strId, out int id))
                    return null;

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                ElementId elemId = new ElementId(id);
#else
                ElementId elemId = new ElementId((long)id);
#endif
                result.Add(SDE_Doc.GetElement(elemId));
            }

            return result;
        }
    }
}
