using Autodesk.Revit.UI;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class MainMenu : System.Windows.Controls.Page, IDockablePaneProvider
    {
        public MainMenu()
        {
            MainMenu_VM = new MVVMCore_MainMenu.MainViewModel();
            
            InitializeComponent();
            DataContext = MainMenu_VM;
        }

        /// <summary>
        /// Ссылка на ViewModel
        /// </summary>
        public MVVMCore_MainMenu.MainViewModel MainMenu_VM {  get; private set; }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Left
            };
        }

        private void SetParamsOpenHoleByIOSElems_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ParamsOpenHoleByIOSElemsForm paramsForm = new ParamsOpenHoleByIOSElemsForm(MainMenu_VM);
            if ((bool)paramsForm.ShowDialog())
            {
                MainMenu_VM = paramsForm.ParamsOpenHoleByIOSElemsForm_VM;
                DataContext = MainMenu_VM;
            }
        }
    }
}
