using KPLN_Library_Forms.UI;
using System.ComponentModel;
using System.Media;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class UserStringInfo : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _userInputName;
        private string _userInputPath;

        public UserStringInfo()
        {
            InitializeComponent();
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            PreviewKeyDown += new KeyEventHandler(HandleEnter);
            SystemSounds.Beep.Play();
        }

        public UserStringInfo(string lastName, string lastPath) : this()
        {
            UserInputName = lastName;
            UserInputPath = lastPath;
        }

        public string UserInputName
        {
            get => _userInputName;
            set
            {
                _userInputName = value;
                OnPropertyChanged();
            }
        }

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
                Close();
        }

        private void HandleEnter(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                Close();
        }

        private void OnOk(object sender, RoutedEventArgs e) => Close();
    }
}
