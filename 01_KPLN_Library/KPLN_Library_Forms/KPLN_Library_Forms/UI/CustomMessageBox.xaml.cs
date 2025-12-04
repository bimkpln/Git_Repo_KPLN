using System.Windows;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    /// <summary>
    /// Окно для вывода сообщения пользователю
    /// </summary>
    public partial class CustomMessageBox : Window
    {
        public CustomMessageBox(string tbHeader, string tblMainContent)
        {
            InitializeComponent();

            this.LabelHeader.Content = tbHeader;
            this.TblMainContent.Text = tblMainContent;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Показать окно с учетом родительского
        /// </summary>
        public void ShowByOwner(Window owner)
        {
            this.Owner = owner;
            this.SetPosition(owner);
            this.Show();
        }

        /// <summary>
        /// Позиционируем окно рядом с главным окном
        /// </summary>
        private void SetPosition(Window owner)
        {
            this.Left = owner.Left + owner.Width;
            this.Top = owner.Top;
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
