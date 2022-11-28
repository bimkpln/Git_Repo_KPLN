﻿using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    public partial class UserVerify : Window
    {
        private string _inputPassword;

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;

        public UserVerify(string description)
        {
            InitializeComponent();

            this.FormDescription.Text = description;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

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
            else if (e.Key == Key.Enter)
            {
                OnRunClick(sender, e);
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(SearchPassword);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_isRun)
            {
                Status = RunStatus.Close;
            }
        }

        private void PasswordText_Changed(object sender, RoutedEventArgs e)
        {
            PasswordBox passwordBox = (PasswordBox)sender;
            _inputPassword = passwordBox.Password;
        }

        private void OnRunClick(object sender, RoutedEventArgs e)
        {
            _isRun = true;

            if (CheckVerify(_inputPassword)) { Status = RunStatus.Run; }
            else { Status = RunStatus.CloseBecauseError; }

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

                if (truePassowrd.Equals(password)) { return true; }

                return false;

            }
            throw new Exception("Отсутсвует файл с паролем. Напиши BIM-менеджеру");
        }
    }
}
