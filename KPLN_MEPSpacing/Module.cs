using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_MEPSpacing
{
    public class Module : IExternalModule
    {
        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            return Result.Succeeded;
        }
    }
}
