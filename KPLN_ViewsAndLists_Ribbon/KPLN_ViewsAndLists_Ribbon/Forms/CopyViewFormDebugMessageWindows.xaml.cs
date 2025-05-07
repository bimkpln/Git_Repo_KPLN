using System;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DebugMessageWindow : Window
    {
        public DebugMessageWindow(string message)
        {
            InitializeComponent();
            UpdateMessageTextBox(message);
        }

        void UpdateMessageTextBox(string fullText)
        {
            MessageTextBox.Document.Blocks.Clear();

            // Разбиваем текст на строки
            var lines = fullText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                var paragraph = new Paragraph(new Run(line));
                paragraph.Margin = new Thickness(0);

                if (line.Contains("ИНФО."))
                {
                    paragraph.Foreground = Brushes.Black;
                }
                else
                {
                    paragraph.Foreground = Brushes.Red;
                }

                MessageTextBox.Document.Blocks.Add(paragraph);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}