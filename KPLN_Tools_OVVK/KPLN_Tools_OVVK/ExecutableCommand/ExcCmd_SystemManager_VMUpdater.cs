using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Forms.Entities;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    /// <summary>
    /// Обновление снимка систем в контексте Revit API.
    /// </summary>
    internal sealed class ExcCmd_SystemManager_VMUpdater : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;

        public ExcCmd_SystemManager_VMUpdater(OVVK_SystemManagerM model)
        {
            _model = model;
        }

        public Result Execute(UIApplication app)
        {
            _model.Refresh();

            return Result.Succeeded;
        }
    }
}
