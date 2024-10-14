using Autodesk.Revit.UI;
using System.Windows;
using System.IO;
using Newtonsoft.Json.Linq;
using Autodesk.Revit.DB.Electrical;
using System.Linq;



namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowСhoice : Window
    {
        public batchAddingParametersWindowСhoice(UIApplication uiapp, string activeFamilyName)
        {        
            InitializeComponent();
            this.uiapp = uiapp;
            this.activeFamilyName = activeFamilyName;
            familyName.Text = activeFamilyName;
        }

        UIApplication uiapp;
        public string activeFamilyName;
        public string paramAction;
        public string paramType;
        public string jsonFileSettingPath;

        // Пакетное добавление общих параметров
        private void Button_NewGeneralParam(object sender, RoutedEventArgs e)
        {
            jsonFileSettingPath = "";

            var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
        }

        // Пакетное добавление параметров семейства
        private void Button_NewFamilyParam(object sender, RoutedEventArgs e)
        {
            var window = new batchAddingParametersWindowFamily(uiapp, activeFamilyName, jsonFileSettingPath);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
        }

        private void Button_LoadParam(object sender, RoutedEventArgs e)
        {          
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Преднастройка (*.json)|*.json";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                jsonFileSettingPath = openFileDialog.FileName;

                string jsonContent = File.ReadAllText(jsonFileSettingPath);
                dynamic jsonFile = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);

                if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["pathFile"] != null && item["groupParameter"] != null && item["nameParameter"] != null && item["instance"] != null && item["grouping"] != null && item["parameterValue"] != null))
                {
                    var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
                    var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                    window.ShowDialog();
                }
                else if (jsonFile is JArray && ((JArray)jsonFile).Any(item => ((string)item["NE"])?.StartsWith("FamilyParamAdd") == true))
                {
                    var window = new batchAddingParametersWindowFamily(uiapp, activeFamilyName, jsonFileSettingPath);
                    var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                    window.ShowDialog();
                }
                else{
                    System.Windows.Forms.MessageBox.Show("Ваш JSON-файл не является файлом преднастроек или повреждён. Пожалуйста, выберите другой файл.", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }

            }
        }
    }
}
