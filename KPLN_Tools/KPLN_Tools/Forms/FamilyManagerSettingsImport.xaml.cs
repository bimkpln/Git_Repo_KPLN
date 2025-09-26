using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class FamilyManagerSettingsImport : Window
    {
        private bool _updating; 

        public FamilyManagerSettingsImport()
        {
            InitializeComponent();

            cbImportParams.Checked += CbImportParams_Checked;

            cbDepartment.Checked += AnyOther_Checked;
            cbFamilyImage.Checked += AnyOther_Checked;
        }

        public bool DoDepartment => cbDepartment.IsChecked == true;
        public bool DoImportParams => cbImportParams.IsChecked == true;
        public bool DoFamilyImage => cbFamilyImage.IsChecked == true;

        private void CbImportParams_Checked(object sender, RoutedEventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                cbDepartment.IsChecked = false;
                cbFamilyImage.IsChecked = false;
            }
            finally { _updating = false; }
        }

        private void AnyOther_Checked(object sender, RoutedEventArgs e)
        {
            if (_updating) return;
            _updating = true;
            try
            {
                cbImportParams.IsChecked = false;
            }
            finally { _updating = false; }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
