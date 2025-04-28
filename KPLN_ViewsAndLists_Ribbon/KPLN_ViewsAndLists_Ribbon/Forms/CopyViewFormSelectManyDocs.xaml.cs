using System;
using System.Windows;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для CopyViewFormSelectManyDocs.xaml
    /// </summary>
    public partial class ManyDocumentsSelectionWindow : Window
    {
        private UIApplication _uiapp;
        private readonly DBRevitDialog[] _dbRevitDialogs;
        private Document _mainDocument;
        private ObservableCollection<FileItem> _items = new ObservableCollection<FileItem>();
        private Dictionary<string, Tuple<string, string, string, View>> _viewOnlyTemplateChanges;

        public ManyDocumentsSelectionWindow(UIApplication uiApp, Document mainDocument, Document additionalDocument, Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges)
        {
            InitializeComponent();

            _uiapp = uiApp;
            _mainDocument = mainDocument;
            _viewOnlyTemplateChanges = viewOnlyTemplateChanges;
            ItemsList.ItemsSource = _items;

            if (additionalDocument != null && additionalDocument.IsValidObject)
            {
                string path = additionalDocument.PathName;

                if (!string.IsNullOrEmpty(path) && Path.GetExtension(path).Equals(".rvt", System.StringComparison.OrdinalIgnoreCase))
                {
                    _items.Add(new FileItem { FullPath = path });
                }
            }
        }

        public class FileItem
        {
            public string FullPath { get; set; }
        }















        // Запуск работы плагина
        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            string debugMessage = "";

            if (_uiapp == null)
            {
                MessageBox.Show($"Приложение Revit не передано", "KPLN. Ошибка");
                return;
            }

            if (_viewOnlyTemplateChanges.Count == 0)
            {
                MessageBox.Show($"Нет изменений для приминения", "KPLN. Информация");
                return;
            }

            _uiapp.DialogBoxShowing += OnDialogBoxShowing;
            _uiapp.Application.FailuresProcessing += OnFailureProcessing;

            string pathMD = _mainDocument.PathName;
            if (!string.IsNullOrEmpty(pathMD))
            {
                _uiapp.OpenAndActivateDocument(pathMD);
            }

            foreach (FileItem item in _items)
            {
                if (File.Exists(item.FullPath))
                {
                    Document doc = null;

                    try
                    {
                        doc = _uiapp.Application.OpenDocumentFile(item.FullPath);
                        debugMessage += $"Открыт документ {doc.Title} ({item.FullPath}).\n";                        
                    }
                    catch (Exception ex)
                    {
                        debugMessage += $"Не удалось открыть документ {item.FullPath}: {ex}.\n";
                        continue;
                    }













                    try
                    {
                        doc.Close();
                        debugMessage += $"Документ закрыт {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                    }
                    catch (Exception ex)
                    {
                        debugMessage += $"ОШИБКА. Не удалось закрыть документ {item.FullPath}: {ex}.\n";
                        continue;
                    }
                }
            }

            _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
            _uiapp.Application.FailuresProcessing -= OnFailureProcessing;

            this.Close();

            DebugMessageWindow debugWindow = new DebugMessageWindow(debugMessage);
            debugWindow.ShowDialog();
        }












        // XAML. Добавить документы
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt",
                Multiselect = true,
                Title = "Выберите файлы Revit"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string fullPath in openFileDialog.FileNames)
                {
                    if (!_items.Any(item => item.FullPath.Equals(fullPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        _items.Add(new FileItem { FullPath = fullPath });
                    }
                }
            }
        }

        // XAML. Удалить документ
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is FileItem item)
            {
                _items.Remove(item);
            }
        }

        // Обработчик ошибок. OnFailureProcessing
        internal void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            FailuresAccessor fa = args.GetFailuresAccessor();
            IList<FailureMessageAccessor> fmas = fa.GetFailureMessages();
            if (fmas.Count > 0)
            {
                List<FailureMessageAccessor> resolveFailures = new List<FailureMessageAccessor>();
                foreach (FailureMessageAccessor fma in fmas)
                {
                    try
                    {
                        fa.DeleteWarning(fma);
                    }
                    catch
                    {
                        fma.SetCurrentResolutionType(
                            fma.HasResolutionOfType(FailureResolutionType.DetachElements)
                            ? FailureResolutionType.DetachElements
                            : FailureResolutionType.DeleteElements);

                        resolveFailures.Add(fma);
                    }
                }

                if (resolveFailures.Count > 0)
                {
                    fa.ResolveFailures(resolveFailures);
                    args.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
                }
            }
        }

        // Обработчик ошибок. OnDialogBoxShowing
        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            if (args.Cancellable)
            {
                args.Cancel();
            }
            else
            {
                DBRevitDialog currentDBDialog = null;
                if (string.IsNullOrEmpty(args.DialogId))
                {
                    TaskDialogShowingEventArgs taskDialogShowingEventArgs = args as TaskDialogShowingEventArgs;
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => !string.IsNullOrEmpty(rd.Message) && taskDialogShowingEventArgs.Message.Contains(rd.Message));
                }
                else
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => !string.IsNullOrEmpty(rd.DialogId) && args.DialogId.Contains(rd.DialogId));

                if (currentDBDialog == null)
                {
                    return;
                }
            }
        }
    }
}
