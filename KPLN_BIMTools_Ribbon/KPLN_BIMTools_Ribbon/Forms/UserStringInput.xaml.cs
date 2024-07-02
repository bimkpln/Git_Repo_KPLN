using System.Media;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class UserStringInput : Window
    {
        private bool _canRunByName;
        private bool _canRunByPathTo;

        public UserStringInput()
        {
            InitializeComponent();

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            PreviewKeyDown += new KeyEventHandler(HandleEnter);
            SystemSounds.Beep.Play();
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        public string UserInputName { get; private set; }
        
        public string UserInputPath { get; private set; }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                IsRun = false;
                Close();
            }
        }

        private void HandleEnter(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.OnOk(sender, e);
                Close();
            }
        }

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

        private void NameChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            UserInputName = textBox.Text;

            if (!string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 5)
                _canRunByName = true;
            else
                _canRunByName = false;

            BtnEnableSwitch();
        }

        private void PathChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            UserInputPath = textBox.Text;

            if (!string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 10)
                _canRunByPathTo = true;
            else
                _canRunByPathTo = false;

            BtnEnableSwitch();
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            if (_canRunByName && _canRunByPathTo)
                btnOk.IsEnabled = true;
            else
                btnOk.IsEnabled = false;
        }
    }
}
