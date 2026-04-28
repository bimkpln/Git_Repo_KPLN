using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_CoordinatorAI.Common
{
    public class StaticAvailable : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            return true;
        }
    }
}