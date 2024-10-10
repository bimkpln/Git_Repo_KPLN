using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using System.ComponentModel;
using System.Media;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class UserPathInputForm : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _canRunByPathTo;
        private string _userInputPath;

        public UserPathInputForm()
        {
            InitializeComponent();
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            PreviewKeyDown += new KeyEventHandler(HandleEnter);
            SystemSounds.Beep.Play();
        }

        public UserPathInputForm(string lastPath) : this()
        {
            UserInputPath = lastPath;
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        public string UserInputPath
        {
            get => _userInputPath;
            set
            {
                _userInputPath = value;
                OnPropertyChanged();
            }
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

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
            if (UserInputPathVerify())
            {
                IsRun = true;
                Close();
            }
            else
            {
                CustomMessageBox cmb = new CustomMessageBox(
                    "Ошибка",
                    $"Путь \"{UserInputPath}\" к файлу/папке не соответвует критериям из описания. Исправь и повтори попытку");
                cmb.ShowDialog();
            }
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            if (_canRunByPathTo)
                btnOk.IsEnabled = true;
            else
                btnOk.IsEnabled = false;
        }

        private void InputName_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tBox = sender as TextBox;

            // Меняю настройки кликабельности кнопок
            if (!string.IsNullOrWhiteSpace(tBox.Text) && tBox.Text.Length > 10)
                _canRunByPathTo = true;
            else
                _canRunByPathTo = false;

            BtnEnableSwitch();
        }

        /// <summary>
        /// Проверка корректности введенного пользователем пути
        /// </summary>
        /// <returns></returns>
        private bool UserInputPathVerify() => UserInputPath.EndsWith(".rvt") || UserInputPath.EndsWith("\\");
    }
}
