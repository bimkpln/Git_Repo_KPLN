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
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };
        }

        private void HistoryFilters_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var history = MainMenu_VM.History;
            var window = new OpeningHistoryFilterWindow(history.CreateFilterSettings());
            if (Module.CurrentUIApplication != null)
                new System.Windows.Interop.WindowInteropHelper(window).Owner = Module.CurrentUIApplication.MainWindowHandle;

            if (window.ShowDialog() == true)
                history.ApplyFilterSettings(window.Settings);
        }

        private void SetParamsOpenHoleByIOSElems_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                ParamsOpenHoleByIOSElemsForm paramsForm = new ParamsOpenHoleByIOSElemsForm(MainMenu_VM);
                if (paramsForm.ShowDialog() == true)
                {
                    MainMenu_VM = paramsForm.ParamsOpenHoleByIOSElemsForm_VM;
                    DataContext = MainMenu_VM;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show("Не удалось открыть настройки отверстий.\n" + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex));
            }
        }
    }
}
