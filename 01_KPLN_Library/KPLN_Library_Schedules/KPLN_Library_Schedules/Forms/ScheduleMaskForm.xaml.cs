using System;
using System.Windows;
using System.Windows.Controls;
using KPLN_Library_Schedules.Common.Schedules;

namespace KPLN_Library_Schedules.Forms
{
    public partial class ScheduleMaskForm : Window
    {
        private readonly ScheduleMaskCatalog _catalog;
        public string Pattern => PatternBox.Text;

        // Сохраняем прежние сигнатуры для внешних потребителей общей библиотеки.
        public ScheduleMaskForm(string pattern, bool forName = false)
            : this(pattern, ScheduleMaskCatalog.CopyByLevels(forName)) { }

        public ScheduleMaskForm(string pattern, bool forName, bool parameterMode)
            : this(pattern, forName, parameterMode, parameterMode) { }

        public ScheduleMaskForm(string pattern, bool forName, bool parameterMode, bool managerMode, bool allowOriginal = true)
            : this(pattern, managerMode ? ScheduleMaskCatalog.Constructor(forName, parameterMode, allowOriginal)
                : ScheduleMaskCatalog.CopyByLevels(forName, allowOriginal)) { }

        public ScheduleMaskForm(string pattern, ScheduleMaskCatalog catalog)
        {
            _catalog = catalog ?? throw new ArgumentNullException(nameof(catalog));
            InitializeComponent();
            ContextText.Text = catalog.Context;
            SampleText.Text = catalog.Sample;
            PatternBox.ToolTip = catalog.Context;
            TokensBox.ItemsSource = catalog.Choices;
            TokensBox.SelectedIndex = catalog.Choices.Count > 0 ? 0 : -1;
            PatternBox.Text = pattern ?? "";
            PatternBox.CaretIndex = PatternBox.Text.Length;
            UpdateExample(null, null);
        }

        private void SelectToken(object sender, SelectionChangedEventArgs args)
        {
            var choice = TokensBox.SelectedItem as ScheduleMaskChoice;
            if (DescriptionText != null)
                DescriptionText.Text = choice?.Description ?? "";
        }

        private void InsertToken(object sender, RoutedEventArgs args)
        {
            var choice = TokensBox.SelectedItem as ScheduleMaskChoice;
            if (choice == null)
                return;

            int position = PatternBox.SelectionStart;
            PatternBox.SelectedText = choice.Token;
            PatternBox.Focus();
            PatternBox.Select(position + choice.Token.Length, 0);
        }

        private void UpdateExample(object sender, TextChangedEventArgs args)
        {
            if (ExampleText == null)
                return;

            try { ExampleText.Text = "Пример результата: " + _catalog.Example(Pattern); }
            catch (ArgumentException ex) { ExampleText.Text = ex.Message; }
        }

        private void Apply(object sender, RoutedEventArgs args)
        {
            try
            {
                _catalog.Example(Pattern);
                DialogResult = true;
            }
            catch (ArgumentException ex) { MessageBox.Show(this, ex.Message, Title); }
        }

        public static void ValidateManagerPattern(string pattern, bool allowValue, bool allowOriginal = true)
        {
            ScheduleMaskCatalog.Constructor(allowValue, allowValue, allowOriginal).Validate(pattern);
        }
    }
}
