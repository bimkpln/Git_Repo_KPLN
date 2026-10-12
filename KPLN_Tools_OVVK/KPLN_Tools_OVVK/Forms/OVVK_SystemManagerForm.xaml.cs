using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_OVVK.ExternalEventHandler;
using KPLN_Tools_OVVK.Forms.ViewModels;
using System;
using System.Windows;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_SystemManagerForm : Window
    {
        private static OVVK_SystemManagerForm _currentInstance;
        private OVVK_SystemManagerSubscriptionHandler _subscriptionHandler;

        public OVVK_SystemManagerForm(Document doc)
        {
            CurrentOVVK_SystemManagerVM = new OVVK_SystemManagerVM(this, doc);

            InitializeComponent();
            DataContext = CurrentOVVK_SystemManagerVM;

            _currentInstance = this;
            Closed += OnClosed;
        }


        public OVVK_SystemManagerVM CurrentOVVK_SystemManagerVM { get; }

        internal void StartTracking(UIApplication app)
        {
            _subscriptionHandler = new OVVK_SystemManagerSubscriptionHandler(
                app, CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM, Close);
            _subscriptionHandler.StartTracking();
        }

        public static bool TryActivateExisting()
        {
            if (_currentInstance == null || !_currentInstance.IsLoaded)
            {
                _currentInstance = null;
                return false;
            }

            if (_currentInstance.WindowState == WindowState.Minimized)
                _currentInstance.WindowState = WindowState.Normal;

            _currentInstance.Show();
            _currentInstance.Activate();
            _currentInstance.Focus();
            return true;
        }


        private void OnClosed(object sender, EventArgs e)
        {
            _subscriptionHandler?.RequestStopTracking();
            CurrentOVVK_SystemManagerVM.Release();

            if (ReferenceEquals(_currentInstance, this))
                _currentInstance = null;
        }
    }
}
