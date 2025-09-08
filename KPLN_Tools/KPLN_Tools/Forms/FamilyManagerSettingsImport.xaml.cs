using System.Windows;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для FamilyManagerSettingsImport.xaml
    /// </summary>
    public partial class FamilyManagerSettingsImport : Window
    {
        public FamilyManagerSettingsImport()
        {
            InitializeComponent();
        }

        public bool DoDepartment => cbDepartment.IsChecked == true;
        public bool DoStage => cbStage.IsChecked == true;
        public bool DoProject => cbProject.IsChecked == true;
        public bool DoImportParams => cbImportParams.IsChecked == true;
        public bool DoFamilyImage => cbFamilyImage.IsChecked == true;

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true; 
        }
    }
}
