using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;


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
        private readonly UIApplication _uiapp;
        private readonly Document _mainDocument;

        private readonly ObservableCollection<FileItem> _items = new ObservableCollection<FileItem>();
        private readonly Dictionary<string, Tuple<string, string, string, string, View>> _viewOnlyTemplateChanges;

        private readonly bool _leaveOpened;

        private string _smallDebugMessage;
        private string _debugMessage;

        public ManyDocumentsSelectionWindow(UIApplication uiApp,Document mainDocument,Dictionary<string, Tuple<string, string, string, string, View>> viewOnlyTemplateChanges, bool leaveOpened, bool useRevitServer)
        {
            InitializeComponent();

            _revitVersion = int.Parse(uiApp.Application.VersionNumber);
            _uiapp = uiApp;
            _mainDocument = mainDocument;

            _viewOnlyTemplateChanges = viewOnlyTemplateChanges;
            _leaveOpened = leaveOpened;

            ItemsList.ItemsSource = _items;

            _debugMessage = "";

            if (useRevitServer)
            {
                AddButton.Click -= AddButton_Click;
                AddButton.Click += AddButton_RevitServer_Click;
            }
        }

        public class FileItem
        {
            public string FullPath { get; set; }

            public string Name
            {
                get => FullPath.Split('\\').Last().TrimEnd(".rvt".ToCharArray());
            }
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
            KPLN_Looker.Module.RunAutoChecks = false;

            List<ModelPath> processedPaths = new List<ModelPath>();
            try
            {

                foreach (FileItem item in _items)
                {
                    bool isWorkShared = false;
                    if (File.Exists(item.FullPath) || item.FullPath.Contains("RSN"))
                    {
                        Document openedDoc = null;
                        UIDocument openeUIdDoc = null;
                        try
                        {
                            ModelPath fileModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(item.FullPath);

                            if (fileModelPath.ServerPath)
                                isWorkShared = true;
                            else
                            {
                                BasicFileInfo basicFileInfo = BasicFileInfo.Extract(item.FullPath);
                                isWorkShared = basicFileInfo.IsWorkshared;
                            }

                            // Определяю имя файла для открытия. Если это ФХ - то нужно открыть локальную копию
                            ModelPath openDocModelPath = fileModelPath;
                            if (isWorkShared)
                            {
                                string localFileName = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Documents\" + $"{item.Name}_{_uiapp.Application.Username}_{DateTime.Now:HHmmss}.rvt");
                                ModelPath localOpenDocModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localFileName);
                                WorksharingUtils.CreateNewLocal(openDocModelPath, localOpenDocModelPath);

                                openDocModelPath = localOpenDocModelPath;
                            }

                            WorksetConfiguration silentWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                            OpenOptions silentOptions = new OpenOptions();
                            silentOptions.SetOpenWorksetsConfiguration(silentWorksetConfig);

                            if (_leaveOpened)
                            {
                                openeUIdDoc = _uiapp.OpenAndActivateDocument(openDocModelPath, silentOptions, false);
                                openedDoc = openeUIdDoc.Document;
                            }
                            else
                                openedDoc = _uiapp.Application.OpenDocumentFile(openDocModelPath, silentOptions);

                            _debugMessage += $"ИНФО. Получен документ {openedDoc.Title} ({item.FullPath}).\n";

                            processedPaths.Add(openDocModelPath);
                        }
                        catch (Exception ex)
                        {
                            _debugMessage += $"ОШИБКА. Не удалось открыть документ {item.FullPath}: {ex.Message}.\n";
                            _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                            continue;
                        }

                        using (Transaction trans = new Transaction(openedDoc, "KPLN. Копирование шаблонов видов"))
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
                                    string statusCreate3D = kvp.Value.Item4;
                                    View templateView = kvp.Value.Item5;

                                    if (statusView == "resave" && templateView != null)
                                    {
                                        View existingTemplate = new FilteredElementCollector(openedDoc)
                                                .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);
                                        List<View> viewsUsingexistingTemplate = null;

                                        View finalTemplateInTarget = null;
                                        ElementId finalTemplateId = ElementId.InvalidElementId;

                                        // Удаление
                                        if (statusCopyView == "ignoreCopyView")
                                        {
                                            if (existingTemplate != null)
                                            {
                                                viewsUsingexistingTemplate = new FilteredElementCollector(openedDoc)
                                                    .OfClass(typeof(View))
                                                    .Cast<View>()
                                                    .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                                    .ToList();

                                                try
                                                {
                                                    string existingTemplateName = existingTemplate.Name;
                                                    existingTemplate.Name = $"{existingTemplateName}_DeleteTemp{DateTime.Now:yyyyMMddHHmmss}";
                                                    openedDoc.Delete(existingTemplate.Id);

                                                    _debugMessage += $"ИНФО. Временный шаблон {existingTemplate.ToString()} удалён.\n";
                                                }
                                                catch (Exception ex)
                                                {
                                                    _debugMessage += $"ОШИБКА. Не удалось удалить шаблон {existingTemplate.ToString()}: {ex.Message}.\n";
                                                    _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                    continue;
                                                }
                                            }

                                            CopyPasteOptions options = new CopyPasteOptions();
                                            options.SetDuplicateTypeNamesHandler(new ReplaceDuplicateTypeNamesHandler());
                                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                                _mainDocument,
                                                new List<ElementId> { templateView.Id },
                                                openedDoc,
                                                null,
                                                options
                                            );





                                            ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                            View copiedTemplateViewNew = openedDoc.GetElement(copiedTemplateIdNew) as View;
                                            string templateName = templateView.Name;

                                            try
                                            {
                                                copiedTemplateViewNew.Name = viewTemplateName;

                                                _debugMessage += $"ИНФО. Шаблон {viewTemplateName} переименован.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                _debugMessage += $"ОШИБКА. Не удалось переименовать шаблон {viewTemplateName}: {ex.Message}\n";
                                                _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                continue;
                                            }

                                            if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                            {
                                                foreach (View view in viewsUsingexistingTemplate)
                                                {
                                                    try
                                                    {
                                                        view.ViewTemplateId = copiedTemplateIdNew;

                                                        _debugMessage += $"ИНФО. Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        _debugMessage += $"ОШИБКА. Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид '{view.Name}': {ex.Message}.\n";
                                                        _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                        continue;
                                                    }
                                                }
                                            }

                                            finalTemplateInTarget = copiedTemplateViewNew;
                                            finalTemplateId = copiedTemplateIdNew;
                                        }
                                        else if (statusCopyView == "copyView")
                                        {
                                            if (existingTemplate != null)
                                            {
                                                viewsUsingexistingTemplate = new FilteredElementCollector(openedDoc)
                                                    .OfClass(typeof(View))
                                                    .Cast<View>()
                                                    .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                                    .ToList();

                                                try
                                                {
                                                    existingTemplate.Name = $"{viewTemplateName}_{DateTime.Now:yyyyMMdd_HHmmss}";

                                                    _debugMessage += $"ИНФО. Шаблон {existingTemplate.Name} переименован.\n";
                                                }
                                                catch (Exception ex)
                                                {
                                                    _debugMessage += $"ОШИБКА. Не удалось переименовать шаблон {existingTemplate.Name}: {ex.Message}\n";
                                                    _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                    continue;
                                                }
                                            }

                                            CopyPasteOptions options = new CopyPasteOptions();
                                            options.SetDuplicateTypeNamesHandler(new ReplaceDuplicateTypeNamesHandler());
                                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                                _mainDocument,
                                                new List<ElementId> { templateView.Id },
                                                openedDoc,
                                                null,
                                                options
                                            );

                                            ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                            View copiedTemplateViewNew = openedDoc.GetElement(copiedTemplateIdNew) as View;

                                            copiedTemplateViewNew.Name = viewTemplateName;

                                            if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                            {
                                                foreach (View view in viewsUsingexistingTemplate)
                                                {
                                                    try
                                                    {
                                                        view.ViewTemplateId = copiedTemplateIdNew;

                                                        _debugMessage += $"ИНФО. Шаблон {copiedTemplateViewNew.Name} назначен на вид {view.Name}.\n";
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        _debugMessage += $"ОШИБКА. Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид '{view.Name}': {ex.Message}.\n";
                                                        _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                                        continue;
                                                    }
                                                }
                                            }

                                            finalTemplateInTarget = copiedTemplateViewNew;
                                            finalTemplateId = copiedTemplateIdNew;
                                        }

                                        if (statusView == "resave" && statusCreate3D == "create3D" && templateView.ViewType == ViewType.ThreeD)
                                        {
                                            try
                                            {
                                                if (finalTemplateInTarget == null || finalTemplateId == ElementId.InvalidElementId)
                                                {
                                                    finalTemplateInTarget = new FilteredElementCollector(openedDoc)
                                                        .OfClass(typeof(View)).Cast<View>()
                                                        .FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);
                                                    finalTemplateId = finalTemplateInTarget?.Id ?? ElementId.InvalidElementId;
                                                }

                                                if (finalTemplateInTarget == null)
                                                    throw new InvalidOperationException($"Не найден 3D-шаблон '{viewTemplateName}' в целевом документе.");

                                                var vft3d = new FilteredElementCollector(openedDoc)
                                                    .OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                                    .FirstOrDefault(t => t.ViewFamily == ViewFamily.ThreeDimensional);
                                                if (vft3d == null)
                                                    throw new InvalidOperationException("Не найден ViewFamilyType для 3D.");

                                                string finalName = MakeUniqueViewName(openedDoc, viewTemplateName);

                                                View3D v3d = View3D.CreateIsometric(openedDoc, vft3d.Id);
                                                v3d.ViewTemplateId = finalTemplateId;
                                                v3d.Name = finalName;

                                                _debugMessage += $"ИНФО. Создан 3D-вид '{v3d.Name}' и применён шаблон '{finalTemplateInTarget.Name}'.\n";
                                            }
                                            catch (Exception ex)
                                            {
                                                _debugMessage += $"ОШИБКА. Не удалось создать 3D-вид из шаблона '{viewTemplateName}': {ex.Message}.\n";
                                                _smallDebugMessage += $"ОШИБКА. {item.FullPath}: 3D из '{viewTemplateName}' — {ex.Message}.\n";
                                            }
                                        }
                                    }
                                }
                            }

                            trans.Commit();
                        }


                        if (_leaveOpened)
                            _smallDebugMessage += $"ИНФО. {item.FullPath}: " +
                                $"Документ обработан и оставлен открытым. " +
                                $"\n!!!ВАЖНО!!! Изменения НЕ сохранены/синхронизированы, подразумевается пользовательские действия по сохранению/синхронизации.\n";
                        else
                        {
                            // Синхронизирую модель из хранилища
                            if (isWorkShared)
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

                                    openedDoc.SynchronizeWithCentral(transOptions, options);

                                    _debugMessage += $"ИНФО. Документ синхронизирован {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                                }
                                catch (Exception ex)
                                {
                                    _debugMessage += $"ОШИБКА. Не удалось синхронизировать документ {item.FullPath}: {ex.Message}.\n";
                                    _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                    continue;
                                }
                            }
                            // Сохраняю обычный файл
                            else
                            {
                                try
                                {
                                    openedDoc.Save();
                                    _debugMessage += $"ИНФО. Документ сохранён {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                                }
                                catch (Exception ex)
                                {
                                    _debugMessage += $"ОШИБКА. Не удалось сохранить документ {item.FullPath}: {ex.Message}.\n";
                                    _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                    continue;
                                }
                            }

                            // Закрываю файл
                            try
                            {
                                openedDoc.Close(false);
                                _debugMessage += $"ИНФО. Документ закрыт {Path.GetFileNameWithoutExtension(item.FullPath)} ({item.FullPath}).\n";
                            }
                            catch (Exception ex)
                            {
                                _debugMessage += $"ОШИБКА. Не удалось закрыть документ {item.FullPath}: {ex.Message}.\n";
                                _smallDebugMessage += $"ОШИБКА. {item.FullPath}: {ex.Message}.\n";
                                continue;
                            }

                            _smallDebugMessage += $"ИНФО. {item.FullPath}: Документ обработан, изменения сохранены/синхронизированы.\n";
                        }
                    }
                    else
                        throw new Exception($"Файл по указанному пути - отсутсвует: {item.FullPath}");
                }

                // Закрываю это окно
                this.Close();

                // Закрываю основное окно
                IntPtr hWnd = FindWindow(null, "KPLN. Копирование шаблонов вида");
                if (hWnd != IntPtr.Zero)
                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);

            }
            catch (Exception ex)
            {
                _debugMessage += $"ОШИБКА. Обратисть к разработчику. Работа остановлена с критической ошибкой: {ex.Message}.\n";
                _smallDebugMessage += $"ОШИБКА. Обратисть к разработчику. Работа остановлена с критической ошибкой: {ex.Message}.\n";
            }
            finally
            {
                _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                _uiapp.Application.FailuresProcessing -= OnFailureProcessing;
                KPLN_Looker.Module.RunAutoChecks = true;
            }

            // Открываю окно с результатами
            DebugMessageWindow debugWindow = new DebugMessageWindow(_uiapp, _smallDebugMessage, _debugMessage, processedPaths);
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

        private void AddButton_RevitServer_Click(object sender, RoutedEventArgs e)
        {

            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
            if (rsFilesPickForm == null)
                return;

            if ((bool)rsFilesPickForm.ShowDialog())
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
                args.Cancel();
            else
            {
                DBRevitDialog currentDBDialog = null;
                if (string.IsNullOrEmpty(args.DialogId))
                {
                    TaskDialogShowingEventArgs taskDialogShowingEventArgs = args as TaskDialogShowingEventArgs;
                    currentDBDialog = DBMainService
                        .DBRevitDialogColl
                        .FirstOrDefault(rd => !string.IsNullOrEmpty(rd.Message) && taskDialogShowingEventArgs.Message.Contains(rd.Message));
                }
                else
                    currentDBDialog = DBMainService
                        .DBRevitDialogColl
                        .FirstOrDefault(rd => !string.IsNullOrEmpty(rd.DialogId) && args.DialogId.Contains(rd.DialogId));

                if (currentDBDialog == null)
                {
                    _smallDebugMessage += $"ОШИБКА. Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека";
                    return;
                }

                if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                {
                    if (args.OverrideResult((int)taskDialogResult))
                        _smallDebugMessage += $"ИНФО. Окно {args.DialogId} успешно закрыто. Была применена команда {currentDBDialog.OverrideResult}";
                    else
                        _smallDebugMessage += $"ОШИБКА. Окно {args.DialogId} не удалось обработать. Была применена команда {currentDBDialog.OverrideResult}, но она не сработала!";
                }
                else
                    _smallDebugMessage += $"Не удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!";
            }
        }

        public class ReplaceDuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
            }
        }

        private static string MakeUniqueViewName(Document doc, string baseName)
        {
            bool NameExists(string n) =>
                new FilteredElementCollector(doc)
                    .OfClass(typeof(View)).Cast<View>()
                    .Any(v => !v.IsTemplate && v.Name.Equals(n, StringComparison.OrdinalIgnoreCase));

            if (!NameExists(baseName))
                return baseName;

            string n1 = $"{baseName} (3D)";
            if (!NameExists(n1))
                return n1;

            int i = 2;
            while (true)
            {
                string cand = $"{baseName} (3D) {i}";
                if (!NameExists(cand))
                    return cand;
                i++;
                if (i > 9999) throw new InvalidOperationException("Не удалось подобрать уникальное имя 3D-вида.");
            }
        }
    }
}