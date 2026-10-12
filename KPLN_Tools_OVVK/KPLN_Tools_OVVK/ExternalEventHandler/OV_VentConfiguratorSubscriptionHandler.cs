using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Tools_OVVK.Services.VentConfigurator;
using KPLN_Tools_OVVK.Forms.ViewModels;

namespace KPLN_Tools_OVVK.ExternalEventHandler
{
    internal sealed class OV_VentConfiguratorSubscriptionHandler
    {
        private UIApplication _application;
        private readonly OV_VentConfiguratorVM _viewModel;
        private readonly VentFamilyService _service;
        private bool _stopRequested;

        internal OV_VentConfiguratorSubscriptionHandler(UIApplication application, OV_VentConfiguratorVM viewModel)
        {
            _application = application;
            _viewModel = viewModel;
            _service = viewModel.FamilyService;
        }

        // Вызывается из ExtCmd в контексте Revit API.
        internal void StartTracking()
        {
            _application.Idling += OnIdling;
            _service.RefreshProjects(_application);
        }

        internal void RequestStopTracking()
        {
            _stopRequested = true;
            _service.Close();
        }

        private void OnIdling(object sender, IdlingEventArgs args)
        {
            // Closed приходит вне API-контекста: снимаем подписку на следующем Idling.
            if (_stopRequested)
            {
                _application.Idling -= OnIdling;
                _application = null;
                _service.ClearAutomaticSnapshots();
                return;
            }

            if (!_viewModel.IsBusy)
                _service.TryUpdateActiveProject(_application, false);
        }
    }
}
