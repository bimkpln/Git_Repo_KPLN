using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.Services.VentConfigurator;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class ExcCmd_OV_VentConfigurator : IExecutableCommand
    {
        private readonly VentFamilyService _service;
        private readonly FamilyRequest _request;

        internal ExcCmd_OV_VentConfigurator(VentFamilyService service, FamilyRequest request)
        {
            _service = service;
            _request = request;
        }

        public Result Execute(UIApplication app)
        {
            if (_service.IsClosed)
                return Result.Cancelled;
            _service.Execute(app, _request);
            return Result.Succeeded;
        }
    }
}
