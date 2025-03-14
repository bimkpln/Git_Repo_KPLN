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
        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;

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
        public RunStatus Status { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Status = RunStatus.Close;
                Close();
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(tBox);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_isRun)
            {
                Status = RunStatus.Close;
            }
        }

        private void OnBtnApply(object sender, RoutedEventArgs e)
        {
            _isRun = true;
            Status = RunStatus.Run;
            DialogResult = true;
            UserInput = tBox.Text;

            Close();
        }
    }
}
