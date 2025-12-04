using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class UserVerify : Window
    {
        private string _inputPassword;

        public UserVerify(string description)
        {
            InitializeComponent();

            this.FormDescription.Text = description;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter)
                OnRunClick(sender, e);
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(SearchPassword);

        private void PasswordText_Changed(object sender, RoutedEventArgs e)
        {
            PasswordBox passwordBox = (PasswordBox)sender;
            _inputPassword = passwordBox.Password;
        }

        private void OnRunClick(object sender, RoutedEventArgs e)
        {
            
            if (CheckVerify(_inputPassword))
                DialogResult = true;
            else 
                DialogResult = false;

            Close();
        }

        private bool CheckVerify(string password)
        {
            FileInfo fi = new FileInfo(@"X:\BIM\5_Scripts\Git_Repo_KPLN\01_KPLN_Library\KPLN_Library_Forms\KPLN_Library_Forms\passwordForBIM.txt");
            if (fi.Exists)
            {
                string truePassowrd = string.Empty;
                using (StreamReader sr = fi.OpenText())
                {
                    while (sr.Peek() >= 0)
                    {
                        truePassowrd = sr.ReadLine();
                    }
                }

                if (truePassowrd.Equals(password)) 
                    return true;

                return false;

            }
            throw new Exception("Отсутсвует файл с паролем. Напиши BIM-менеджеру");
        }
    }
}
