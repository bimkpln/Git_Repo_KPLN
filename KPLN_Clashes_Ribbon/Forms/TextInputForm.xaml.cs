using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для TextInput.xaml
    /// </summary>
    public partial class TextInputForm : Window
    {
        public bool IsConfirmed = false;
        
        public TextInputForm(Window parent, string header)
        {
            Owner = parent;
            
            InitializeComponent();
            
            tbHeader.Text = header;
            runBtn.IsEnabled = false;
        }

        /// <summary>
        /// Текст, введенный пользователем
        /// </summary>
        public string UserComment { get; private set; }

        private void OnBtnApply(object sender, RoutedEventArgs e)
        {
            UserComment = tbox.Text;
            IsConfirmed = true;
            Close();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(tbox);
        }

        private void TBX_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            runBtn.IsEnabled = VerifyInput(textBox.Text);
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!IsAllowedCharacter(c))
                {
                    // Запретить ввод символа
                    e.Handled = true;
                    break;
                }
            }
        }

        /// <summary>
        /// Проверка символов ввода
        /// </summary>
        /// <param name="c">Символ</param>
        /// <returns></returns>
        private bool IsAllowedCharacter(char c)
        {
            // Проверка, разрешен ли символ (кириллица и дефис)
            return char.IsLetter(c)
                || char.IsWhiteSpace(c)
                || char.IsDigit(c)
                || c == '-'
                || c == '?'
                || c == ','
                || c == '.'
                || c == '_';
        }

        /// <summary>
        /// Проверка на ввод данных
        /// </summary>
        private bool VerifyInput(string msg) => !string.IsNullOrWhiteSpace(msg) && tbox.Text.Length > 5;
    }
}
