using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
using System.ComponentModel;

namespace KPLN_Tools_OVVK.ExternalEventHandler
{
    internal sealed class OVVK_SystemManagerSubscriptionHandler
    {
        private UIApplication _application;
        private readonly OVVK_SystemManagerM _model;
        private readonly Action _closeWindow;
        private bool _refreshPending;
        private bool _stopRequested;
        private bool _isTracking;

        internal OVVK_SystemManagerSubscriptionHandler(UIApplication application, OVVK_SystemManagerM model, Action closeWindow)
        {
            _application = application;
            _model = model;
            _closeWindow = closeWindow;
        }


        internal void StartTracking()
        {
            if (_isTracking || _stopRequested)
                return;

            _model.UpdateSelectedRevitViews(_application.ActiveUIDocument);
            _application.Idling += OnIdling;
            try
            {
                _application.Application.DocumentChanged += OnDocumentChanged;
                _model.PropertyChanged += OnModelPropertyChanged;
                _isTracking = true;
            }
            catch
            {
                _application.Idling -= OnIdling;
                throw;
            }
        }

        internal void RequestStopTracking()
        {
            _stopRequested = true;
        }

        private void OnDocumentChanged(object sender, DocumentChangedEventArgs args)
        {
            if (!_stopRequested && _model.CurrentDoc.IsValidObject && _model.CurrentDoc.Equals(args.GetDocument()))
                _refreshPending = true;
        }

        private void OnModelPropertyChanged(object sender, PropertyChangedEventArgs args)
        {
            // Команды сами обновляют модель после транзакции: повторный обход на Idling не нужен.
            if (args.PropertyName == nameof(OVVK_SystemManagerM.RefreshRevision))
                _refreshPending = false;
        }

        private void OnIdling(object sender, IdlingEventArgs args)
        {
            // Closed вызывается из WPF: снимаем подписки на следующем событии Revit.
            if (_stopRequested)
            {
                StopTracking();
                return;
            }

            if (!_model.CurrentDoc.IsValidObject)
            {
                // Документ уже закрыт; отписываемся здесь, пока есть контекст Revit API.
                _stopRequested = true;
                StopTracking();
                _closeWindow();
                return;
            }

            _model.IsDocumentActive = _application.ActiveUIDocument != null
                && _model.CurrentDoc.Equals(_application.ActiveUIDocument.Document);
            // Выбор в диспетчере проекта отслеживается и в Revit 2020, без SelectionChanged.
            if (!_model.IsBusy)
                _model.UpdateSelectedRevitViews(_application.ActiveUIDocument, _refreshPending);
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

        private void StopTracking()
        {
            if (!_isTracking)
                return;

            _application.Idling -= OnIdling;
            _application.Application.DocumentChanged -= OnDocumentChanged;
            _model.PropertyChanged -= OnModelPropertyChanged;
            _application = null;
            _isTracking = false;
        }
    }
}
