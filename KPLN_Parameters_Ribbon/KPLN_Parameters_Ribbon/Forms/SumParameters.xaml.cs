using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Parameters_Ribbon.ExternalEventHandler;
using KPLN_Parameters_Ribbon.Forms.Common;
using KPLN_Parameters_Ribbon.Forms.ViewModels;
using System.Windows;

namespace KPLN_Parameters_Ribbon.Forms
{
    public partial class SumParameters : Window
    {
        private static SumParameters _currentInstance;

#if !Debug2020 && !Revit2020
        private readonly ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;
#endif

        public SumParameters(UIApplication uiapp)
        {
            CurrentSumParametersVM = new SumParametersVM(uiapp);

            InitializeComponent();

            DataContext = CurrentSumParametersVM;

#if !Debug2020 && !Revit2020
            BtnUpdate.Visibility = Visibility.Collapsed;

            _selExtEv = FormEventSubscriptionHelper.CreateSelectionChangedEvent(handler =>
            {
                _selHandler = handler;
                _selHandler.CurrentSumParametersVM = CurrentSumParametersVM;
            });

            ExternalEvent unsubSelExtEv = FormEventSubscriptionHelper.CreateSelectionUnsubscribeEvent(OnSelectionChanged);
            FormEventSubscriptionHelper.SubscribeSelectionChanged(uiapp, this, OnSelectionChanged, unsubSelExtEv);
#endif

            _currentInstance = this;
            Closed += (sender, args) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        public SumParametersVM CurrentSumParametersVM { get; set; }

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e) => _selExtEv?.Raise();
#endif

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
    }
}
