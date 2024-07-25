using System.ComponentModel;
using System.Media;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class UserStringInput : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _canRunByName;
        private bool _canRunByPathTo;
        private string _userInputName;
        private string _userInputPath;

        public UserStringInput()
        {
            InitializeComponent();
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            PreviewKeyDown += new KeyEventHandler(HandleEnter);
            SystemSounds.Beep.Play();
        }

        public UserStringInput(string lastName, string lastPath) : this()
        {
            UserInputName = lastName;
            UserInputPath = lastPath;
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        public string UserInputName
        {
            get => _userInputName;
            set
            {
                if (value != _userInputName)
                {
                    _userInputName = value;
                    OnPropertyChanged();
                    // Меняю настройки кликабельности кнопок
                    if (!string.IsNullOrWhiteSpace(_userInputName) && _userInputName.Length > 5)
                        _canRunByName = true;
                    else
                        _canRunByName = false;

                    BtnEnableSwitch();
                }
            }
        }

        public string UserInputPath
        {
            get => _userInputPath;
            set
            {
                if (value != _userInputPath)
                {
                    _userInputPath = value;
                    OnPropertyChanged();
                    // Меняю настройки кликабельности кнопок
                    if (!string.IsNullOrWhiteSpace(_userInputPath) && _userInputPath.Length > 10)
                        _canRunByPathTo = true;
                    else
                        _canRunByPathTo = false;

                    BtnEnableSwitch();
                }
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
            IsRun = true;
            Close();
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
