using Autodesk.Revit.UI;
using KPLN_MEPBender.Forms.ViewModels;
using System.Windows;

namespace KPLN_MEPBender.Forms
{
    public partial class MepBenderForm : Window
    {
        private static MepBenderForm _currentInstance;

        public MepBenderForm(UIApplication uiapp)
        {
            CurrentMepBenderVM = new MepBenderVM(this, uiapp);

            InitializeComponent();
            DataContext = CurrentMepBenderVM;

            _currentInstance = this;
            Closed += (sender, args) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        public MepBenderVM CurrentMepBenderVM { get; set; }

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
