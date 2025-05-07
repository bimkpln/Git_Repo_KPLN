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
using System.Runtime.InteropServices;
using System.Threading;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для CopyViewFormSelectManyDocs.xaml
    /// </summary>
    public partial class ManyDocumentsSelectionWindow : Window
    {
        private UIApplication _uiapp;      
        private Document _mainDocument;

        private ObservableCollection<FileItem> _items = new ObservableCollection<FileItem>();
        private Dictionary<string, Tuple<string, string, string, View>> _viewOnlyTemplateChanges;

        string debugMessage;

        public ManyDocumentsSelectionWindow(UIApplication uiApp, Document mainDocument, Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges)
        {
            InitializeComponent();

            _uiapp = uiApp;
            _mainDocument = mainDocument;

            _viewOnlyTemplateChanges = viewOnlyTemplateChanges;

            ItemsList.ItemsSource = _items;

            debugMessage = "";
        }

        public class FileItem
        {
            public string FullPath { get; set; }
        }

        private void ProcessButton_Click(object sender, RoutedEventArgs e)
        {
            if (_uiapp == null)
            {
                MessageBox.Show($"Revit-поток не передан или не обработан в плагине. Пожалуйста, перезагрузите Revit.", "KPLN. Ошибка");
                return;
            }

            if (_items.Count == 0)
            {
                MessageBox.Show($"Нет выбранных документов для внесения изменений", "KPLN. Информация");
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

                        debugMessage += $"ИНФО. Открыт документ {doc.Title} ({item.FullPath}).\n";                        
                    }
                    catch (Exception ex)
                    {
                        debugMessage += $"ОШИБКА. Не удалось открыть документ {item.FullPath}: {ex}.\n";

                        continue;
                    }

                    using (Transaction trans = new Transaction(doc, "KPLN. Копирование шаблонов видов"))
                    {
                        trans.Start();

                        // Шаблоны видов
                        if (_viewOnlyTemplateChanges.Count > 0)
                        {
                            foreach (var kvp in _viewOnlyTemplateChanges)
                            {
                                string viewTemplateName = kvp.Key;
                                string statusView = kvp.Value.Item1;
                                string statusResaveInView = kvp.Value.Item2;
                                string statusCopyView = kvp.Value.Item3;
                                View templateView = kvp.Value.Item4;

                                if (statusView == "resave" && templateView != null)
                                {
                                    View existingTemplate = new FilteredElementCollector(doc)
                                            .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);
                                    List<View> viewsUsingexistingTemplate = null;

                                    // Удаление
                                    if (statusCopyView == "ignoreCopyView")
                                    {                                       
                                        if (existingTemplate != null)
                                        {
                                            viewsUsingexistingTemplate = new FilteredElementCollector(doc)
                                                .OfClass(typeof(View))
                                                .Cast<View>()
                                                .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                                .ToList();

                                            try
                                            {
                                                string existingTemplateName = existingTemplate.Name;
                                                existingTemplate.Name = $"{existingTemplateName}_DeleteTemp{DateTime.Now:yyyyMMddHHmmss}";
                                                doc.Delete(existingTemplate.Id);

                                                debugMessage += $"ИНФО. Временный шаблон {existingTemplate.ToString()} удалён.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось удалить шаблон {existingTemplate.ToString()}: {ex.Message}.\n";
                                            }
                                        }                                        
                                                                             
                                        CopyPasteOptions options = new CopyPasteOptions();
                                        options.SetDuplicateTypeNamesHandler(new ReplaceDuplicateTypeNamesHandler());
                                        ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                            _mainDocument,
                                            new List<ElementId> { templateView.Id },
                                            doc,
                                            null,
                                            options
                                        );                                       

                                        ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                        View copiedTemplateViewNew = doc.GetElement(copiedTemplateIdNew) as View;
                                        string templateName = templateView.Name;

                                        try
                                        {
                                            copiedTemplateViewNew.Name = viewTemplateName;

                                            debugMessage += $"ИНФО. Шаблон {viewTemplateName} переименован.\n";
                                        }
                                        catch (Exception ex)
                                        {
                                            debugMessage += $"ОШИБКА. Не удалось переименовать шаблон {viewTemplateName}: {ex.Message}\n";
                                        }

                                        if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                        {
                                            foreach (View view in viewsUsingexistingTemplate)
                                            {
                                                try
                                                {
                                                    view.ViewTemplateId = copiedTemplateIdNew;

                                                    debugMessage += $"ИНФО. Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
                                                }
                                                catch (Exception ex)
                                                {
                                                    debugMessage += $"ОШИБКА. Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид '{view.Name}': {ex.Message}.\n";
                                                }
                                            }
                                        }
                                    }
                                    else if (statusCopyView == "copyView")
                                    {
                                        if (existingTemplate != null)
                                        {
                                            viewsUsingexistingTemplate = new FilteredElementCollector(doc)
                                                .OfClass(typeof(View))
                                                .Cast<View>()
                                                .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                                .ToList();

                                            try
                                            {
                                                existingTemplate.Name = $"{viewTemplateName}_{DateTime.Now:yyyyMMdd_HHmmss}";

                                                debugMessage += $"ИНФО. Шаблон {existingTemplate.Name} переименован.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось переименовать шаблон {existingTemplate.Name}: {ex.Message}\n";
                                            }
                                        }
                                                                             
                                        CopyPasteOptions options = new CopyPasteOptions();
                                        options.SetDuplicateTypeNamesHandler(new ReplaceDuplicateTypeNamesHandler());
                                        ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                            _mainDocument,
                                            new List<ElementId> { templateView.Id },
                                            doc,
                                            null,
                                            options
                                        );
                                        
                                        ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                        View copiedTemplateViewNew = doc.GetElement(copiedTemplateIdNew) as View;

                                        copiedTemplateViewNew.Name = viewTemplateName;

                                        if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                        {
                                            foreach (View view in viewsUsingexistingTemplate)
                                            {
                                                try
                                                {
                                                    view.ViewTemplateId = copiedTemplateIdNew;

                                                    debugMessage += $"ИНФО. Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
                                                }
                                                catch (Exception ex)
                                                {
                                                    debugMessage += $"ОШИБКА. Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид '{view.Name}': {ex.Message}.\n";
                                                }
                                            }
                                        }                                        
                                    }                              
                                }
                            }
                        }

                        trans.Commit();
                    }
                    try
                    {
                        doc.Save();

                        debugMessage += $"ИНФО. Документ сохранён {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                    }
                    catch (Exception ex)
                    {
                        debugMessage += $"ОШИБКА. Не удалось сохранить документ {item.FullPath}: {ex}.\n";
                        continue;
                    }

                    if (doc.IsWorkshared)
                    {
                        try
                        {
                            var options = new SynchronizeWithCentralOptions();
                            var transOptions = new TransactWithCentralOptions();
                            var relinquishOptions = new RelinquishOptions(true)
                            {
                                StandardWorksets = true,
                                FamilyWorksets = true,
                                ViewWorksets = true,
                                UserWorksets = true
                            };

                            options.SetRelinquishOptions(relinquishOptions);
                            options.Comment = "Автоматическая синхронизация через скрипт";

                            doc.SynchronizeWithCentral(transOptions, options);

                            debugMessage += $"ИНФО. Документ синхронизирован {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                        }
                        catch (Exception ex)
                        {
                            debugMessage += $"ОШИБКА. Не удалось синхронизировать документ {item.FullPath}: {ex}.\n";
                            continue;
                        }
                    }

                    try
                    {
                        doc.Close(false);

                        debugMessage += $"ИНФО. Документ закрыт {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
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

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is FileItem item)
            {
                _items.Remove(item);
            }
        }

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

        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            if (args.Cancellable)
            {
                args.Cancel();
            }
            else
            {                
                return;              
            }
        }

        public class ReplaceDuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}