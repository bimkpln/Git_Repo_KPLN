using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using KPLN_ViewsAndLists_Ribbon.Forms.MVVM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class ListRenumberForm : Window
    {
        private readonly UIApplication _uiapp;
        private readonly Document _document;
        private readonly ApiHandler _handler;
        private readonly ExternalEvent _externalEvent;
        private ListRenumberPreviewForm _preview;
        private ListRenumberPreviewVM _previewVM;
        private bool _closed;
        private bool _dirty = true;
        private bool _applying;
        private bool _resourcesReleased;

        public ListRenumberForm(UIApplication uiapp, IEnumerable<ViewSheet> sortedSheets, ParameterSet tbs)
        {
            _uiapp = uiapp;
            _document = uiapp.ActiveUIDocument.Document;
            InitializeComponent();
            new WindowInteropHelper(this).Owner = uiapp.MainWindowHandle;

            var parameters = tbs.Cast<Parameter>().OrderBy(p => p.Definition.Name).ToArray();
            var vm = new ListRenumberVM(uiapp, sortedSheets, parameters);
            DataContext = vm;
            _handler = new ApiHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            vm.RequestApiAction = action => Request(action);
            vm.PreviewAction = preview =>
            {
                _previewVM = preview;
                _preview = new ListRenumberPreviewForm(preview);
                new WindowInteropHelper(_preview).Owner = uiapp.MainWindowHandle;
                _preview.ApplyAction = () => Request(() =>
                {
                    if (_previewVM.TryApply()) Close();
                });
                _preview.Closed += (sender, args) =>
                {
                    _handler.Action = null;
                    _preview = null;
                    _previewVM = null;
                    if (!_closed) { IsEnabled = true; Show(); Activate(); }
                };
                Hide();
                _preview.Show();
                return null;
            };

            uiapp.Application.DocumentChanged += DocumentChanged;
            uiapp.Idling += OnIdling;
            uiapp.ApplicationClosing += OnApplicationClosing;
            Closed += OnClosed;
        }

        private void Request(Action action)
        {
            if (_closed || _handler.Action != null) return;
            _handler.Action = action;
            IsEnabled = false;
            _previewVM?.SetBusy(true);
            var result = _externalEvent.Raise();
            if (result != ExternalEventRequest.Accepted && result != ExternalEventRequest.Pending)
            {
                _handler.Action = null;
                IsEnabled = true;
                _previewVM?.SetBusy(false);
                MessageBox.Show("Revit пока не принял запрос. Повторите действие после завершения текущей команды.");
            }
        }

        private void DocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            if (_closed) return;

            if (e.GetDocument().Equals(_document))
            {
                _dirty = true;
                _previewVM?.SetContextError("Модель изменилась. Ожидается повторная проверка.");
            }
        }

        private bool CheckDocument(UIApplication app)
        {
            if (!_document.IsValidObject) { Close(); return false; }
            bool active = app.ActiveUIDocument != null && app.ActiveUIDocument.Document.Equals(_document);
            _previewVM?.SetContextError(active ? null : "Активируйте проект, из которого открыта нумерация.");
            return active;
        }

        private void OnIdling(object sender, IdlingEventArgs e)
        {
            if (_applying) return;
            if (_closed)
            {
                ReleaseApiResources();
                return;
            }
            try
            {
                bool active = CheckDocument(_uiapp);
                IsEnabled = active && _handler.Action == null;
                if (active && _dirty && _previewVM != null)
                {
                    _previewVM.RefreshModel();
                    _dirty = false;
                }
            }
            catch (Exception ex)
            {
                _previewVM?.SetContextError("Не удалось проверить модель: " + ex.Message);
            }
            finally
            {
                if (_closed) ReleaseApiResources();
            }
        }

        private void OnClosed(object sender, EventArgs e)
        {
            _closed = true;
            _handler.Action = null;
            _preview?.Close();
            // Closed вызывается WPF вне API-контекста. Очистку выполнит Revit.
        }

        private void OnApplicationClosing(object sender, ApplicationClosingEventArgs e)
        {
            try
            {
                Close();
            }
            finally
            {
                ReleaseApiResources();
            }
        }

        // Вызывать только из Idling, Execute или ApplicationClosing.
        private void ReleaseApiResources()
        {
            if (_resourcesReleased) return;

            _uiapp.Application.DocumentChanged -= DocumentChanged;
            _uiapp.Idling -= OnIdling;
            _uiapp.ApplicationClosing -= OnApplicationClosing;
            _handler.Action = null;
            _externalEvent.Dispose();
            _resourcesReleased = true;
        }

        private sealed class ApiHandler : IExternalEventHandler
        {
            private readonly ListRenumberForm _owner;
            internal Action Action { get; set; }
            internal ApiHandler(ListRenumberForm owner) { _owner = owner; }
            public string GetName() => "KPLN: Предпросмотр нумерации листов";

            public void Execute(UIApplication app)
            {
                var action = Action;
                Action = null;
                if (_owner._closed)
                {
                    _owner.ReleaseApiResources();
                    return;
                }
                if (action == null) return;
                _owner._applying = true;
                try
                {
                    if (_owner.CheckDocument(app)) action();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Нумерация листов");
                }
                finally
                {
                    _owner._applying = false;
                    _owner.IsEnabled = true;
                    _owner._previewVM?.SetBusy(false);
                    if (_owner._closed) _owner.ReleaseApiResources();
                }
            }
        }
    }
}
