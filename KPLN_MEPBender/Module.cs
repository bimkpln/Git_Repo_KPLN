using Autodesk.Revit.UI;
using KPLN_Loader.Common;

namespace KPLN_MEPBender
{
    public class Module : IExternalModule
    {
        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Кнопка добавляется в KPLN_TrailingMEP.Module. Так сделано, чтобы сгенерить стэк из кнопок

            return Result.Succeeded;
        }
    }
}
