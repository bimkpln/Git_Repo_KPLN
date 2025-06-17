using System;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DebugMessageWindow : Window
    {
        string _message;

        public DebugMessageWindow(string smallMessage, string message)
        {
            InitializeComponent();
            UpdateMessageTextBox(smallMessage);
            _message = message;
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

        private void SaveMessageButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            string fileName = $"debug_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

            saveFileDialog.FileName = fileName;
            saveFileDialog.DefaultExt = ".txt";
            saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, _message);
                    System.Windows.MessageBox.Show("Файл успешно сохранён.", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    
    }
}