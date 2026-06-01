using System.Windows;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class BitrixCoordinatorDialog : Window
    {
        public BitrixCoordinatorDialog()
        {
            InitializeComponent();
        }

        public string UserId
        {
            get { return (UserIdTextBox.Text ?? string.Empty).Trim(); }
        }

        public string UserName
        {
            get { return (UserNameTextBox.Text ?? string.Empty).Trim(); }
        }

        private void OnAddClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(UserId) || string.IsNullOrWhiteSpace(UserName))
            {
                MessageBox.Show(this, "Заполните ID и имя пользователя.", "Bitrix24", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}