using System.Media;
using System.Windows;

namespace KPLN_Library_Forms.UI
{
    public partial class UserDialog : Window
    {
        public UserDialog(string header, string body)
        {
            InitializeComponent();

            tbHeader.Text = header;
            tbBody.Text = body;

            SystemSounds.Beep.Play();
        }

        public UserDialog(string header, string body, string footer) : this (header, body)
        {
            tbFooter.Text = footer;
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            IsRun = false;
            Close();
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            Close();
        }
    }
}
