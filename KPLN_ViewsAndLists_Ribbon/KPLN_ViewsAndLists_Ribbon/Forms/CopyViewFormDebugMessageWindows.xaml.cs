using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DebugMessageWindow : Window
    {
        private readonly List<ModelPath> _paths;
        string _message;

        private readonly OpenBatchHandler _batchHandler;
        private readonly ExternalEvent _batchEvent;

        public DebugMessageWindow(UIApplication uiApp, string smallMessage, string message, List<ModelPath> paths)
        {
            InitializeComponent();
            _paths = paths ?? new List<ModelPath>();

            UpdateMessageTextBox(smallMessage);
            _message = message;

            _batchHandler = new OpenBatchHandler(uiApp);
            _batchEvent = ExternalEvent.Create(_batchHandler);
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

        private void OpenDocButton_Click(object sender, RoutedEventArgs e)
        {
            if (_paths == null || _paths.Count == 0)
            {
                System.Windows.MessageBox.Show("Список пуст.", "Инфо",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                _batchHandler.SetBatch(_paths);
                _batchEvent.Raise();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка запуска открытия: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            this.Close();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }        
    }

    public class OpenBatchHandler : IExternalEventHandler
    {
        private readonly UIApplication _uiapp;
        private readonly Queue<ModelPath> _queue = new Queue<ModelPath>();

        public OpenBatchHandler(UIApplication uiapp) { _uiapp = uiapp; }

        public void SetBatch(IEnumerable<ModelPath> paths)
        {
            _queue.Clear();
            if (paths == null) return;
            foreach (var p in paths)
                if (p != null) _queue.Enqueue(p);
        }

        public void Execute(UIApplication app)
        {
            var appDoc = _uiapp.Application;

            var wsCfg = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            var opts = new OpenOptions
            {
                DetachFromCentralOption = DetachFromCentralOption.DoNotDetach,
                Audit = false
            };
            opts.SetOpenWorksetsConfiguration(wsCfg);

            while (_queue.Count > 0)
            {
                var mp = _queue.Dequeue();
                try
                {
                    _uiapp.OpenAndActivateDocument(mp, opts, false);
                }
                catch (Exception ex)
                {
                    string visible = "";
                    try { visible = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp); } catch { }
                    TaskDialog.Show("Открытие документа",
                        $"Не удалось открыть:\n{visible}\n\n{ex.Message}");
                }
            }
        }

        public string GetName() => "KPLN.OpenBatchHandler";
    }



}