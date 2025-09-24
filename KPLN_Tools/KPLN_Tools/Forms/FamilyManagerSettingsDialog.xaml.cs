using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class SettingsDialog : Window
    {
        public int Result { get; private set; } = 0;

        public SettingsDialog()
        {
            InitializeComponent();
        }

        private void BtnUpdateFileBD_Click(object sender, RoutedEventArgs e)
        {
            Result = 1;
            DialogResult = true; 
        }

        private void BtnUpdateSettingBD_Click(object sender, RoutedEventArgs e)
        {
            Result = 2;
            DialogResult = true;
        }

        private void BtnDebug_Click(object sender, RoutedEventArgs e)
        {
            Result = 3;
            DialogResult = true; 
        }
    }
}