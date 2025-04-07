using Autodesk.Revit.UI;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class MainMenu : System.Windows.Controls.Page, IDockablePaneProvider
    {
        public MainMenu()
        {
            InitializeComponent();
            DataContext = new MVVMCore_MainMenu.ViewModel();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Left
            };
        }
    }
}
