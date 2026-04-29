using System;
using System.Windows;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class SheetFilterWindow : Window
    {
        public SheetFilterWindow(string parameterName, string defaultValue)
        {
            InitializeComponent();

            PromptTextBlock.Text = string.Format(
                "Выбрать листы, где параметр '{0}' содержит:",
                parameterName);

            FilterTextBox.Text = defaultValue ?? string.Empty;
        }

        public string FilterValue { get; private set; }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FilterTextBox.Focus();
            FilterTextBox.SelectAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string value = (FilterTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(value))
            {
                MessageBox.Show(
                    this,
                    "Введите значение фильтра листов.",
                    Title,
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            FilterValue = value;
            DialogResult = true;
            Close();
        }
    }
}
