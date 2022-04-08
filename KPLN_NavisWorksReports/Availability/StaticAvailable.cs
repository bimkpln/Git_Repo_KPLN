using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_NavisWorksReports.Availability
{
    public class StaticAvailable : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            return true;
        }
    }
}
