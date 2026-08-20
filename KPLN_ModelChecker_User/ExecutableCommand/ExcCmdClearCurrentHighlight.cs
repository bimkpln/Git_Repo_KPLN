using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_User.Common;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class ExcCmdClearCurrentHighlight : IExecutableCommand
    {
        public Result Execute(UIApplication app)
        {
            CheckListAnnotationsService.ClearCurrentHighlight(app);

            return Result.Succeeded;
        }
    }
}
