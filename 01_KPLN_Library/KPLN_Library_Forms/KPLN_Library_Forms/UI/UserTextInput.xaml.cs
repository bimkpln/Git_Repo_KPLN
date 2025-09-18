using System;
using System.Windows;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    /// <summary>
    /// Логика взаимодействия для TextInput.xaml
    /// </summary>
    public partial class UserTextInput : Window
    {
        public UserTextInput(string tbHeader)
        {
            InitializeComponent();

            this.tbHeader.Text = tbHeader;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Сообщение, которое ввел пользователь
        /// </summary>
        public string UserInput { get; private set; }

        /// <summary>
        /// Статус запуска
        /// </summary>
        [Obsolete("Нужно использовать DialogResult")]
        public RunStatus Status { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Status = RunStatus.Close;
                DialogResult = false;
                Close();
            }

            if (e.Key == Key.Enter)
            {
                DialogResult = true;
                Close();
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(tBox);

        private void OnBtnApply(object sender, RoutedEventArgs e)
        {
            Status = RunStatus.Run;
            
            DialogResult = true;
            UserInput = tBox.Text;

            Close();
        }
    }
}
