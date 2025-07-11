using System;
using System.Windows;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Collections.Generic;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
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
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        const uint WM_CLOSE = 0x0010;

        private readonly int _revitVersion;
        private UIApplication _uiapp;      
        private Document _mainDocument;

        private ObservableCollection<FileItem> _items = new ObservableCollection<FileItem>();
        private Dictionary<string, Tuple<string, string, string, View>> _viewOnlyTemplateChanges;

        private bool _openDocument;
        List<string> _worksetPrefixName;
        private bool _useRevitServer;

        string smallDebugMessage;
        string debugMessage;

        public ManyDocumentsSelectionWindow(UIApplication uiApp, Document mainDocument, Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges, bool openDocument, List<string> worksetPrefixName, bool useRevitServer)
        {
            InitializeComponent();

            _revitVersion = int.Parse(uiApp.Application.VersionNumber);
            _uiapp = uiApp;
            _mainDocument = mainDocument;

            _viewOnlyTemplateChanges = viewOnlyTemplateChanges;
            _openDocument = openDocument;
            _worksetPrefixName = worksetPrefixName;
            _useRevitServer = useRevitServer;

            ItemsList.ItemsSource = _items;

            debugMessage = "";

            if (useRevitServer)
            {
                AddButton.Click -= AddButton_Click;
                AddButton.Click += AddButton_RevitServer_Click;
            }
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
                try
                {
                    ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(pathMD);

                    WorksetConfiguration silentWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                    OpenOptions silentOptions = new OpenOptions
                    {
                        DetachFromCentralOption = DetachFromCentralOption.DoNotDetach,
                        Audit = false
                    };
                    silentOptions.SetOpenWorksetsConfiguration(silentWorksetConfig);

                    UIDocument uiDoc = _uiapp.OpenAndActivateDocument(modelPath, silentOptions, false);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("ОШИБКА", $"{ex.Message}");
                }
            }

            foreach (FileItem item in _items)
            {
                if (File.Exists(item.FullPath) || item.FullPath.Contains("RSN"))
                {
                    Document doc = null;

                    try
                    {
                        bool isRevitServerPath = item.FullPath.StartsWith("RSN:", StringComparison.OrdinalIgnoreCase);
                        ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(item.FullPath);

                        WorksetConfiguration silentWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                        OpenOptions silentOptions = new OpenOptions
                        {
                            DetachFromCentralOption = DetachFromCentralOption.DoNotDetach,
                            Audit = false
                        };
                        silentOptions.SetOpenWorksetsConfiguration(silentWorksetConfig);
                       
                        if (_openDocument)
                        {
                            UIDocument uiDoc = _uiapp.OpenAndActivateDocument(modelPath, silentOptions, false);
                            doc = uiDoc.Document;
                        }
                        else
                        {
                            doc = _uiapp.Application.OpenDocumentFile(modelPath, silentOptions);
                        }
                              
                        debugMessage += $"ИНФО. Получен документ {doc.Title} ({item.FullPath}).\n";               
                    }
                    catch (Exception ex)
                    {
                        debugMessage += $"ОШИБКА. Не удалось открыть документ {item.FullPath}: {ex.Message}.\n";
                        smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
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
                                                smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                continue;
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
                                            smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                            continue;
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
                                                    smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                    continue;
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
                                                smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                continue;
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
                                                    smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                    continue;
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
                        debugMessage += $"ОШИБКА. Не удалось сохранить документ {item.FullPath}: {ex.Message}.\n";
                        smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
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
                            debugMessage += $"ОШИБКА. Не удалось синхронизировать документ {item.FullPath}: {ex.Message}.\n";
                            smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                            continue;
                        }
                    }

                    if (!_openDocument)
                    {
                        try
                        {
                            doc.Close(false);
                            debugMessage += $"ИНФО. Документ закрыт {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                        }
                        catch (Exception ex)
                        {
                            debugMessage += $"ОШИБКА. Не удалось закрыть документ {item.FullPath}: {ex.Message}.\n";
                            smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                            continue;
                        }
                    }
                }

                smallDebugMessage += $"ИНФО. Документ обработан - {item.FullPath}.\n";
            }
           
            _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
            _uiapp.Application.FailuresProcessing -= OnFailureProcessing;

            this.Close();

            IntPtr hWnd = FindWindow(null, "KPLN. Копирование шаблонов вида");
            if (hWnd != IntPtr.Zero)
            {
                PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }

            DebugMessageWindow debugWindow = new DebugMessageWindow(smallDebugMessage, debugMessage);
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

        private void AddButton_RevitServer_Click (object sender, RoutedEventArgs e)
        {

            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
            if (rsFilesPickForm == null)
                return;

            if ((bool) rsFilesPickForm.ShowDialog())
            {
                foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
                {
                    _items.Add(new FileItem { FullPath = $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}" });
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