using Autodesk.Revit.UI;
using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.Forms.Entities;
using KPLN_Tools_OVVK.Forms.ViewModels;
using KPLN_Tools_OVVK.ExternalEventHandler;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OV_VentConfiguratorForm : Window
    {
        private static OV_VentConfiguratorForm _currentInstance;
        private OV_VentConfiguratorSubscriptionHandler _subscriptionHandler;
        private bool _released;

        internal OV_VentConfiguratorVM CurrentOV_VentConfiguratorVM
        { get; }

        private OV_VentConfiguratorM Model
        {
            get
            {
                return CurrentOV_VentConfiguratorVM.CurrentOV_VentConfiguratorM;
            }
        }

        public OV_VentConfiguratorForm()
        {
            CurrentOV_VentConfiguratorVM = new OV_VentConfiguratorVM(this);
            InitializeComponent();
            DataContext = CurrentOV_VentConfiguratorVM;
            FamilyTypesListBox.ItemsSource = Model.Types;
            CurrentOV_VentConfiguratorVM.SetBusy(false);
            _currentInstance = this;
        }

        internal void StartTracking(UIApplication app)
        {
            _subscriptionHandler = new OV_VentConfiguratorSubscriptionHandler(app, CurrentOV_VentConfiguratorVM);
            try
            {
                _subscriptionHandler.StartTracking();
            }
            catch
            {
                ReleaseResources();
                throw;
            }
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
            if (!_currentInstance.IsVisible)
                _currentInstance.Show();
            _currentInstance.Activate();
            _currentInstance.Focus();
            return true;
        }

        internal object SelectedType
        {
            get
            {
                return FamilyTypesListBox.SelectedItem;
            }

            set
            {
                FamilyTypesListBox.SelectedItem = value;
            }
        }

        internal int SelectedTab
        {
            get
            {
                return ConfigurationTabs.SelectedIndex;
            }

            set
            {
                ConfigurationTabs.SelectedIndex = value;
            }
        }

        internal string TypeName
        {
            get
            {
                return TypeNameTextBox.Text;
            }

            set
            {
                TypeNameTextBox.Text = value;
            }
        }

        internal void ShowConfiguration(bool visible)
        {
            ConfigurationTabs.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            NoTypeTextBlock.Visibility = visible ? Visibility.Collapsed : Visibility.Visible;
        }

        internal void ClearConfiguration()
        {
            TypeNameTextBox.Clear();
            SectionBlocksPanel.Children.Clear();
            SelectedSectionPanel.Children.Clear();
            InstallationWidthEditor.Content = null;
            AdditionalInstallationEditor.Content = null;
            InfoPanel.DataContext = null;
            SectionCountComboBox.ItemsSource = null;
            AddValveCheckBox.IsChecked = false;
            _parameterEditors.Clear();
            _installationLengthText = null;
            _sectionParametersPanel = null;
            ShowConfiguration(false);
        }

        private void FamilyTypesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Model.IsSynchronizing || Model.IsBusy)
                return;
            var selected = FamilyTypesListBox.SelectedItem as FamilyTypeItem;
            if (selected != null)
                CurrentOV_VentConfiguratorVM.SelectType(selected);
        }

        private void TypeName_TextChanged(object sender, TextChangedEventArgs e)
        {
            CurrentOV_VentConfiguratorVM.RenameType(TypeNameTextBox.Text);
        }

        private void SectionCount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SectionCountComboBox.SelectedItem is int count)
                CurrentOV_VentConfiguratorVM.ResizeSections(count);
        }

        private void AddValve_CheckedChanged(object sender, RoutedEventArgs e)
        {
            CurrentOV_VentConfiguratorVM.SetValve(AddValveCheckBox.IsChecked == true);
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!Model.IsBusy)
                return;
            e.Cancel = true;
            CurrentOV_VentConfiguratorVM.SetStatus("Дождитесь завершения операции. Если Revit занят другой командой, завершите её.", false);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            ReleaseResources();
        }

        internal void ReleaseResources()
        {
            if (_released)
                return;
            _released = true;
            CurrentOV_VentConfiguratorVM.Dispose();
            _subscriptionHandler?.RequestStopTracking();
            if (ReferenceEquals(_currentInstance, this))
                _currentInstance = null;
        }
    }
}
