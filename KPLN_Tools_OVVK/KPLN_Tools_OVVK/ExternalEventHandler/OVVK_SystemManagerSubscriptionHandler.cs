using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools_OVVK.Forms.Entities;
using System;

namespace KPLN_Tools_OVVK.ExternalEventHandler
{
    internal sealed class OVVK_SystemManagerSubscriptionHandler
    {
        private UIApplication _application;
        private readonly OVVK_SystemManagerM _model;
        private readonly Action _closeWindow;
        private bool _refreshPending;
        private bool _stopRequested;

        internal OVVK_SystemManagerSubscriptionHandler(UIApplication application, OVVK_SystemManagerM model, Action closeWindow)
        {
            _application = application;
            _model = model;
            _closeWindow = closeWindow;
        }


        internal void StartTracking()
        {
            _application.Idling += OnIdling;
            _application.Application.DocumentChanged += OnDocumentChanged;
        }

        internal void RequestStopTracking()
        {
            _stopRequested = true;
        }

        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (!_stopRequested && _model.CurrentDoc.Equals(args.GetDocument()))
                _refreshPending = true;
        }

        private void OnIdling(object sender, IdlingEventArgs args)
        {
            // Closed вызывается из WPF: снимаем подписки на следующем событии Revit.
            if (_stopRequested)
            {
                _application.Idling -= OnIdling;
                _application.Application.DocumentChanged -= OnDocumentChanged;
                _application = null;
                return;
            }

            if (!_model.CurrentDoc.IsValidObject)
            {
                _closeWindow();
                return;
            }

            _model.IsDocumentActive = _application.ActiveUIDocument != null
                && _model.CurrentDoc.Equals(_application.ActiveUIDocument.Document);
            if (!_refreshPending || !_model.CanRefresh)
                return;


            _refreshPending = false;
            try
            {
                _model.Refresh();
            }
            catch (Exception ex)
            {
                _model.SetOperationMessage("Не удалось обновить список систем: " + ex.Message);
                HtmlOutput.PrintError(ex);
            }
        }
    }
}
