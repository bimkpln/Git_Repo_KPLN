using System.Windows;

namespace KPLN_Tools.Forms
{
    public enum SettingsAction
    {
        None = 0,
        UpdateDb = 1,
        AddDelCategory = 2
    }

    public partial class SettingsWindowNodeManager : Window
    {
        public SettingsAction SelectedAction { get; private set; } = SettingsAction.None;

        public SettingsWindowNodeManager()
        {
            InitializeComponent();
        }

        private void BtnUpdateDb_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = SettingsAction.UpdateDb;
            DialogResult = true;
            Close();
        }

        private void BtnAddDelCategory_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = SettingsAction.AddDelCategory;
            DialogResult = true;
            Close();
        }
    }
}
