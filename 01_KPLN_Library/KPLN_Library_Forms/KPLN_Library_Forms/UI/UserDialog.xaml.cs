using System;
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
                IsRun = false;
                DialogResult = false;
                Close();
            }

            // Обработка Enter - может вызывть запуск последней команды, и окно опять появиться. Запрещено её добавлять
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        [Obsolete("Нужно использовать DialogResult")]
        public bool IsRun { get; private set; }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            IsRun = false;
            Close();
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            DialogResult = true;
            Close();
        }
    }
}
