using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_OVVK.ExternalEventHandler;
using KPLN_Tools_OVVK.Forms.ViewModels;
using System;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using KPLN_Tools_OVVK.Forms.Entities;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_SystemManagerForm : Window
    {
        private static OVVK_SystemManagerForm _currentInstance;
        private OVVK_SystemManagerSubscriptionHandler _subscriptionHandler;
        private bool _syncingSelection;

        public OVVK_SystemManagerForm(Document doc)
        {
            CurrentOVVK_SystemManagerVM = new OVVK_SystemManagerVM(this, doc);

            InitializeComponent();
            DataContext = CurrentOVVK_SystemManagerVM;
            CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM.PropertyChanged += OnModelPropertyChanged;
            SyncSelection();

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


        private void SystemsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_syncingSelection || SystemsList == null)
                return;

            _syncingSelection = true;
            try
            {
                CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM.SetSelectedRows(
                    SystemsList.SelectedItems.Cast<OVVK_SystemManagerRowM>().ToArray());
            }
            finally { _syncingSelection = false; }
        }

        private void OnModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(OVVK_SystemManagerM.SystemRows))
                SyncSelection();
        }

        private void SyncSelection()
        {
            var model = CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM;
            OVVK_SystemManagerRowM[] selected = model.SelectedSystemRows;
            _syncingSelection = true;
            try
            {
                // SelectedItems учитывает и строки вне видимой области при Ctrl+A / Shift.
                SystemsList.ItemsSource = model.SystemRows;
                SystemsList.SelectedItems.Clear();
                foreach (OVVK_SystemManagerRowM row in selected)
                    SystemsList.SelectedItems.Add(row);
                model.SetSelectedRows(selected);
            }
            finally { _syncingSelection = false; }
        }

        private void OnClosed(object sender, EventArgs e)
        {
            CurrentOVVK_SystemManagerVM.CurrentOVVK_SystemManagerM.PropertyChanged -= OnModelPropertyChanged;
            _subscriptionHandler?.RequestStopTracking();
            CurrentOVVK_SystemManagerVM.Release();

            if (ReferenceEquals(_currentInstance, this))
                _currentInstance = null;
        }
    }
}
