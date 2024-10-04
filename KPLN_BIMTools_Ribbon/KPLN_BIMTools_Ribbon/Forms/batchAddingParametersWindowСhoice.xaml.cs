using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;


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

        private void Button_NewParam(object sender, RoutedEventArgs e)
        {
            jsonFileSettingPath = "";

            OpenBatchAddingParametersWindow(activeFamilyName, jsonFileSettingPath);
        }

        private void Button_LoadParam(object sender, RoutedEventArgs e)
        {
           
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Преднастройка (*.json)|*.json";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                jsonFileSettingPath = openFileDialog.FileName;

                OpenBatchAddingParametersWindow(activeFamilyName, jsonFileSettingPath);
            }
        }

        private void OpenBatchAddingParametersWindow(string activeFamilyName, string jsonFileSettingPath)
        {        
            if (generalParam.IsChecked == true) 
            {
                Close();
                var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
                var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                window.ShowDialog();
            } if (familyParam.IsChecked == true) {
                // Окно параметры семейства
            }
        }
    }
}
