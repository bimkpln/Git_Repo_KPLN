using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common.HolesManager
{
    /// <summary>
    /// Класс-контейнер для объединения данных по связям
    /// </summary>
    public class ARLinkData
    {
        public ARLinkData(RevitLinkInstance currentLink)
        {
            CurrentLink = currentLink;
        }
        
        public RevitLinkInstance CurrentLink { get; private set; }

        public List<FilteredElementCollector> CurrentFEC { get; private set; }

        /// <summary>
        /// Подготовить коллекцию FilteredElementCollector 
        /// </summary>
        public void SetFEC(List<BuiltInCategory> bicColl)
        {
            List<FilteredElementCollector> result = new List<FilteredElementCollector>();
            
            foreach (BuiltInCategory bic in bicColl)
            {
                FilteredElementCollector fecIOS = new FilteredElementCollector(CurrentLink.GetLinkDocument())
                    .OfCategory(bic)
                    .WhereElementIsNotElementType();
                
                if (fecIOS.Count() != 0)
                    result.Add(fecIOS);
            }

            CurrentFEC = result;
        }
    }
}
