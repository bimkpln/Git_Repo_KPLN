using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace KPLN_Quantificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для OutputForm.xaml
    /// </summary>
    public partial class OutputForm : Window
    {
        private static OutputForm _instance;

        public static OutputForm GetInstance()
        {
            if (_instance == null)
                _instance = new OutputForm();
            return _instance;
        }

        private OutputForm()
        {
            InitializeComponent();
        }

        public void AddHeaderTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 10;
            margin.Bottom = 5;
            text_block.Margin = margin;
            text_block.FontSize = 15;
            text_block.FontWeight = FontWeights.Black;
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }

        public void AddTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 2;
            margin.Bottom = 2;
            text_block.Margin = margin;
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }

        public void AddErrorTextBlock(string message)
        {
            TextBlock text_block = new TextBlock();
            text_block.Text = message;
            Thickness margin = new Thickness();
            margin.Left = 5;
            margin.Right = 5;
            margin.Top = 2;
            margin.Bottom = 2;
            text_block.Margin = margin;
            text_block.Foreground = new SolidColorBrush(Color.FromArgb(255, 190, 0, 0));
            text_block.Background = new SolidColorBrush(Color.FromArgb(255, 255, 196, 193));
            text_block.TextWrapping = TextWrapping.Wrap;
            text_block.BaselineOffset = 2;
            this.stack_panel.Children.Add(text_block);
        }
        public void AddSuccessTextBlock(string message, string boldKey = null, int? boldTotalGroups = null, int? boldIosGroups = null)
        {
            var text_block = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                BaselineOffset = 2,
                Margin = new Thickness(5, 2, 5, 2),
                Foreground = new SolidColorBrush(Color.FromArgb(255, 60, 140, 0)),
                Background = new SolidColorBrush(Color.FromArgb(255, 203, 255, 174))
            };

            if (boldKey == null && boldTotalGroups == null && boldIosGroups == null)
            {
                text_block.Text = message;
                this.stack_panel.Children.Add(text_block);
                return;
            }

            void AddPlain(string s)
            {
                if (!string.IsNullOrEmpty(s))
                    text_block.Inlines.Add(new Run(s));
            }

            void AddBold(string s)
            {
                if (!string.IsNullOrEmpty(s))
                    text_block.Inlines.Add(new Run(s) { FontWeight = FontWeights.Bold });
            }

            int cursor = 0;
            void BoldFirst(string target)
            {
                if (string.IsNullOrEmpty(target)) return;

                int idx = message.IndexOf(target, cursor, StringComparison.Ordinal);
                if (idx < 0)
                    return;

                AddPlain(message.Substring(cursor, idx - cursor));
                AddBold(message.Substring(idx, target.Length));
                cursor = idx + target.Length;
            }

            if (!string.IsNullOrEmpty(boldKey))
                BoldFirst(boldKey);

            if (boldTotalGroups.HasValue)
            {
                string groupsExact = $"(Групп {boldTotalGroups.Value})";
                int saveCursor = cursor;

                int idxExact = message.IndexOf(groupsExact, cursor, StringComparison.Ordinal);
                if (idxExact >= 0)
                {
                    AddPlain(message.Substring(cursor, idxExact - cursor) + "(Групп ");
                    AddBold(boldTotalGroups.Value.ToString());
                    AddPlain(")");
                    cursor = idxExact + groupsExact.Length;
                }
                else
                {
                    cursor = saveCursor;
                    BoldFirst(boldTotalGroups.Value.ToString());
                }
            }

            if (boldIosGroups.HasValue)
            {
                string iosExact = $"(Групп {boldIosGroups.Value})";
                int saveCursor = cursor;

                int idxExact = message.IndexOf(iosExact, cursor, StringComparison.Ordinal);
                if (idxExact >= 0)
                {
                    AddPlain(message.Substring(cursor, idxExact - cursor) + "(Групп ");
                    AddBold(boldIosGroups.Value.ToString());
                    AddPlain(")");
                    cursor = idxExact + iosExact.Length;
                }
                else
                {
                    cursor = saveCursor;
                    BoldFirst(boldIosGroups.Value.ToString());
                }
            }

            if (cursor < message.Length)
                AddPlain(message.Substring(cursor));

            this.stack_panel.Children.Add(text_block);
        }

        private void On_Closed(object sender, System.ComponentModel.CancelEventArgs e) => _instance = null;
    }
}
