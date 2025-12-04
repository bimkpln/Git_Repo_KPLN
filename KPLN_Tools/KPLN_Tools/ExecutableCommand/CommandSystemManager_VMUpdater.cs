using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools.Forms.Models;

namespace KPLN_Tools.ExecutableCommand
{
    /// <summary>
    /// Класс по обновлению данных для OVVK_SystemManager_ViewModel (кнопки на OnIdling)
    /// </summary>
    internal class CommandSystemManager_VMUpdater : IExecutableCommand
    {
        private readonly OVVK_SystemManager_VM _viewModel;

        public CommandSystemManager_VMUpdater(OVVK_SystemManager_VM viewModel)
        {
            _viewModel = viewModel;
        }

        public Result Execute(UIApplication app)
        {
            _viewModel.UpdateElementColl();
            _viewModel.UpdateSystemParamData();

            return Result.Succeeded;
        }
    }
}
