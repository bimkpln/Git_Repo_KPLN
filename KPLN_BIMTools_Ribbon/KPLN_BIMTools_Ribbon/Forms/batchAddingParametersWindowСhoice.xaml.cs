using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;

namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowСhoice : Window
    {
        public batchAddingParametersWindowСhoice(UIApplication uiapp)
        {
            InitializeComponent();
            this.uiapp = uiapp;
        }

        UIApplication uiapp;
        public string paramAction;
        public string paramType;
        public string jsonFileSettingPath;

        private void Button_NewParam(object sender, RoutedEventArgs e)
        {
            paramAction = "new";
            jsonFileSettingPath = "";

            OpenBatchAddingParametersWindow(paramAction, jsonFileSettingPath);
        }

        private void Button_LoadParam(object sender, RoutedEventArgs e)
        {
            paramAction = "load";

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Преднастройка (*.json)|*.json";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                jsonFileSettingPath = openFileDialog.FileName;

                OpenBatchAddingParametersWindow(paramAction, jsonFileSettingPath);
            }
        }

        private void OpenBatchAddingParametersWindow(string paramAction, string jsonFileSettingPath)
        {        
            if (generalParam.IsChecked == true) 
            {
                Close();
                var window = new batchAddingParametersWindowGeneral(uiapp, paramAction, jsonFileSettingPath);
                var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                window.ShowDialog();
            } if (familyParam.IsChecked == true) { 
                // Окно параметры семейства
            }
    }
    }
}
