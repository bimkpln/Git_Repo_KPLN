using System.Windows;
using System.Windows.Input;

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

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }

            // Обработка Enter - может вызывть запуск последней команды, и окно опять появиться. Запрещено её добавлять
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(tBox);

        private void OnBtnApply(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            UserInput = tBox.Text;

            Close();
        }
    }
}
