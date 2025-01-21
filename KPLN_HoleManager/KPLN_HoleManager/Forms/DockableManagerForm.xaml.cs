using Autodesk.Revit.UI;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_HoleManager.Forms
{
    public partial class DockableManagerForm : Page, IDockablePaneProvider
    {
        public DockableManagerForm()
        {
            InitializeComponent();
            DataContext = new ButtonDataViewModel();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this as FrameworkElement;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };
        }
    }

    // Передаём данные статусов в названия кнопок
    public class ButtonDataViewModel : INotifyPropertyChanged
    {
        private string _approvedButtonText = $"✔️ Утверждено: {null}";
        private string _warningButtonText = $"⚠️ Предупреждения: {null}";
        private string _errorButtonText = $"❌ Ошибки: {null}";

        public string ApprovedButtonText
        {
            get => _approvedButtonText;
            set
            {
                _approvedButtonText = value;
                OnPropertyChanged(nameof(ApprovedButtonText));
            }
        }

        public string WarningButtonText
        {
            get => _warningButtonText;
            set
            {
                _warningButtonText = value;
                OnPropertyChanged(nameof(WarningButtonText));
            }
        }

        public string ErrorButtonText
        {
            get => _errorButtonText;
            set
            {
                _errorButtonText = value;
                OnPropertyChanged(nameof(ErrorButtonText));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
