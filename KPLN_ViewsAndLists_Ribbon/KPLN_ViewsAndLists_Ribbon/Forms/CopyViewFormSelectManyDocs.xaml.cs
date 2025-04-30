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
using System.Windows.Media;
using System.Windows.Documents;

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
        private bool _replaceTypes;

        public ManyDocumentsSelectionWindow(UIApplication uiApp, Document mainDocument, Document additionalDocument, Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges, bool replaceTypes)
        {
            InitializeComponent();

            _uiapp = uiApp;
            _mainDocument = mainDocument;
            _viewOnlyTemplateChanges = viewOnlyTemplateChanges;
            _replaceTypes = replaceTypes;
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
                MessageBox.Show($"Revit-uiapp не передан", "KPLN. Ошибка");
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
                                        if (_replaceTypes)
                                        {
                                            try
                                            {
                                                RemoveDuplicateTypes(_mainDocument, doc, existingTemplate);
                                                debugMessage += $"Связные типы из {templateView.Name}.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось удалить cвязные типы из {templateView.Name}: {ex.Message}.\n";
                                            }
                                        }

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

                                                debugMessage += $"Временный шаблон {existingTemplate.ToString()} удалён.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось удалить шаблон {existingTemplate.ToString()}: {ex.Message}.\n";
                                            }
                                        }                                        

                                        CopyPasteOptions options = new CopyPasteOptions();
                                        options.SetDuplicateTypeNamesHandler(new MyDuplicateTypeNamesHandler());

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

                                            debugMessage += $"Шаблон {viewTemplateName} переименован.\n";
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

                                                    debugMessage += $"Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
                                                }
                                                catch (Exception ex)
                                                {
                                                    debugMessage += $"ОШИБКА. Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид '{view.Name}': {ex.Message}.\n";
                                                }
                                            }
                                        }
                                    }

                                    // Резервная копия
                                    else if (statusCopyView == "copyView")
                                    {
                                        if (_replaceTypes)
                                        {
                                            try
                                            {
                                                RemoveDuplicateTypes(_mainDocument, doc, existingTemplate);
                                                debugMessage += $"Связные типы из {templateView.Name}.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось удалить cвязные типы из {templateView.Name}: {ex.Message}.\n";
                                            }
                                        }

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

                                                debugMessage += $"Шаблон {existingTemplate.Name} переименован.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                debugMessage += $"ОШИБКА. Не удалось переименовать шаблон {existingTemplate.Name}: {ex.Message}\n";
                                            }
                                        }

                                        CopyPasteOptions options = new CopyPasteOptions();
                                        options.SetDuplicateTypeNamesHandler(new MyDuplicateTypeNamesHandler());

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

                                                    debugMessage += $"Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
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

                        debugMessage += $"Документ сохранён {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
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

                            debugMessage += $"Документ синхронизирован {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
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

        // Поиск одинаковых типов в разных документах
        public static HashSet<ElementId> GetUsedTypeIdsFromView(View view)
        {
            Document doc = view.Document;
            HashSet<ElementId> result = new HashSet<ElementId>();

            foreach (Parameter param in view.Parameters)
            {
                if (param.StorageType == StorageType.ElementId)
                {
                    ElementId id = param.AsElementId();
                    if (id != ElementId.InvalidElementId)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem is ElementType)
                        {
                            result.Add(id);
                        }
                    }
                }
            }

            foreach (Category category in doc.Settings.Categories)
            {
                OverrideGraphicSettings ogs = view.GetCategoryOverrides(category.Id);
                if (ogs != null)
                {
                    AddElementTypesFromOverrideGraphicSettings(doc, ogs, result);
                }
            }

            foreach (ElementId filterId in view.GetFilters())
            {
                OverrideGraphicSettings ogs = view.GetFilterOverrides(filterId);
                if (ogs != null)
                {
                    AddElementTypesFromOverrideGraphicSettings(doc, ogs, result);
                }
            }

            return result;
        }

        // Вспомогательный метод для поиск одинаковых типов в разных документах
        private static void AddElementTypesFromOverrideGraphicSettings(Document doc, OverrideGraphicSettings ogs, HashSet<ElementId> result)
        {
            ElementId[] ids =
            {
                ogs.CutLinePatternId,
                ogs.ProjectionLinePatternId,
                ogs.SurfaceForegroundPatternId,
                ogs.SurfaceBackgroundPatternId,
                ogs.CutForegroundPatternId,
                ogs.CutBackgroundPatternId
            };

            foreach (ElementId id in ids)
            {
                if (id != ElementId.InvalidElementId)
                {
                    Element elem = doc.GetElement(id);
                    if (elem is ElementType)
                    {
                        result.Add(id);
                    }
                }
            }
        }

        // Удалить из additionalDocument те типы, у которых такие же имена, как в mainDocument
        void RemoveDuplicateTypes(Document sourceDoc, Document targetDoc, View sourceView)
        {
            var usedTypeIds = GetUsedTypeIdsFromView(sourceView);

            foreach (ElementId typeId in usedTypeIds)
            {
                Element typeInSource = sourceDoc.GetElement(typeId);
                if (typeInSource is ElementType sourceType)
                {
                    ElementType matchingType = new FilteredElementCollector(targetDoc)
                        .OfClass(sourceType.GetType())
                        .Cast<ElementType>()
                        .FirstOrDefault(t => t.Name == sourceType.Name);

                    if (matchingType != null)
                    {
                        try
                        {
                            targetDoc.Delete(matchingType.Id);
                        }
                        catch
                        {

                        }
                    }
                }
            }
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

        // Обработчик копировани типов. Оставляем старые типы, которые уже есть
        public class MyDuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }
    }
}
