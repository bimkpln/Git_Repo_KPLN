﻿using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms.Common;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Loader.Forms
{
    /// <summary>
    /// Interaction logic for LoginForm.xaml
    /// </summary>
    public partial class LoginForm : Window
    {
        public LoginForm(IEnumerable<SubDepartment> subDepartments, bool isExrtraNet)
        {
            InitializeComponent();

            this.Version.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.cbxDepartment.ItemsSource = subDepartments;
            CreatedWPFUser = new WPFUser();

            if (isExrtraNet)
                tbxCompany.IsEnabled = true;
            else
                CreatedWPFUser.Company = "KPLN";

            DataContext = CreatedWPFUser;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);

        }

        public WPFUser CreatedWPFUser { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                OnRunClick(sender, e);
        }

        private void OnRunClick(object sender, RoutedEventArgs e)
        {
            if (runBtn.IsEnabled)
            {
                this.DialogResult = true;
                Close();
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateRunBtnEnabled();
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

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRunBtnEnabled();
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
                && Regex.IsMatch(c.ToString(), @"\p{IsCyrillic}")
                || char.IsWhiteSpace(c) 
                || c == '-';
        }

        /// <summary>
        /// Проверка на ввод данных
        /// </summary>
        private void UpdateRunBtnEnabled()
        {
            // Проверяем, что все TextBox и ComboBox заполнены
            bool allControlsFilled = !string.IsNullOrWhiteSpace(tbxSurname.Text)
                && tbxSurname.Text.Length > 2
                && !string.IsNullOrWhiteSpace(tbxName.Text)
                && tbxName.Text.Length > 2
                && cbxDepartment.SelectedItem != null;

            // Обновляем состояние кнопки на основе результата проверки
            runBtn.IsEnabled = allControlsFilled;
        }
    }
}
