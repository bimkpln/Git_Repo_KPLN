using System.Windows;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class UserDialog : Window
    {
        public UserDialog(string header, string body)
        {
            InitializeComponent();

            tbHeader.Text = header;
            tbBody.Text = body;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public UserDialog(string header, string body, string footer) : this(header, body)
        {
            tbFooter.Text = footer;
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }

            // Обработка Enter - может вызывть запуск последней команды, и окно опять появиться. Запрещено её добавлять
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
