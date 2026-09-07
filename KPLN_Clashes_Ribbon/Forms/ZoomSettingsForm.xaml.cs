using Autodesk.Revit.DB;
using KPLN_Clashes_Ribbon.Forms.ViewModels;
using System.Windows;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ZoomSettingsForm.xaml
    /// </summary>
    public partial class ZoomSettingsForm : Window
    {
        private static ZoomSettingsForm _currentInstance;

        public ZoomSettingsForm(Document doc)
        {
            CurrentZoomSettingsVM = new ZoomSettingsVM(this, doc);

            InitializeComponent();

            DataContext = CurrentZoomSettingsVM;

            _currentInstance = this;
            Closed += (sender, args) =>
            {
                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        public ZoomSettingsVM CurrentZoomSettingsVM { get; set; }

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
