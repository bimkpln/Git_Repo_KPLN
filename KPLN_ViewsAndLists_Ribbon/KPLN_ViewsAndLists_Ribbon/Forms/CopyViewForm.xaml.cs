using System;
using System.IO;
using System.Linq;
using System.Windows.Media;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.DB.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;

using RevitApp = Autodesk.Revit.ApplicationServices.Application;
using Document = Autodesk.Revit.DB.Document;
using Grid = System.Windows.Controls.Grid;


namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public class CopyViewExternalHandler : IExternalEventHandler
    {
        public Action<UIApplication> ExecuteAction;

        public void Execute(UIApplication app)
        {
            ExecuteAction?.Invoke(app);
        }

        public string GetName() => "KPLN. ExternalHandler";
    }

    public partial class CopyViewForm : Window
    {
        private readonly ExternalEvent _copyEvent;
        private readonly CopyViewExternalHandler _copyHandler;

        UIApplication _uiapp;
        private readonly DBRevitDialog[] _dbRevitDialogs;

        Document mainDocument = null;
        Document additionalDocument = null;

        Document heritageMainDocument = null;
        Document heritageAdditionalDocument = null;

        SortedDictionary<string, View> mainTemplates;
        private Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, View>>();

        public CopyViewForm(UIApplication uiapp, Document heritageMainDocument, Document heritageAdditionalDocument)
        {
            InitializeComponent();
            _uiapp = uiapp;

            if (heritageMainDocument == null)
            {
                mainDocument = uiapp.ActiveUIDocument.Document;
            }
            else
            {
                mainDocument = heritageMainDocument;
            }

            if (heritageAdditionalDocument != null)
            {
                additionalDocument = heritageAdditionalDocument;

                if (additionalDocument != null)
                {
                    string additionalDocName = additionalDocument.Title;

                    StackPanel panel = new StackPanel
                    {
                        Orientation = Orientation.Vertical,
                        Margin = new Thickness(10, 4, 0, 4)
                    };

                    TextBlock textBlock = new TextBlock
                    {
                        Text = "Имя открытого документа",
                        Height = 15,
                        FontSize = 11
                    };

                    System.Windows.Controls.TextBox textBox = new System.Windows.Controls.TextBox
                    {
                        Name = "TB_OtherDocumentName",
                        Height = 25,
                        Width = 655,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        Margin = new Thickness(0, 0, 15, 0),
                        Padding = new Thickness(3),
                        Cursor = System.Windows.Input.Cursors.Arrow,
                        IsReadOnly = true,
                        Text = additionalDocName,
                        ToolTip = additionalDocument.PathName
                    };

                    panel.Children.Add(textBlock);
                    panel.Children.Add(textBox);
                    SP_OtherFile.Children.Add(panel);

                    // Обработка документа
                    if (mainDocument != null && additionalDocument != null)
                    {
                        BTN_OpenTempalte.IsEnabled = true;

                        DisplayAllTemplatesInPanel(mainDocument, additionalDocument);
                        viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, View>>();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
                        BTN_OpenTempalte.IsEnabled = false;
                        CHK_ReplaceTypes.IsEnabled = false;
                        CHK_ReplaceTypes.IsChecked = false;
                        CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                        BTN_RunOneFile.IsEnabled = false;
                        BTN_RunManyFile.IsEnabled = false;
                    }
                }
            }

            string mainDocPath = mainDocument.PathName;
            string mainDocName = mainDocument.Title;

            if (mainDocPath != null && mainDocName != null)
            {
                TB_ActivDocumentName.Text = $"{mainDocName}";
                TB_ActivDocumentName.ToolTip = $"{mainDocPath}";
            }

            _copyHandler = new CopyViewExternalHandler();
            _copyEvent = ExternalEvent.Create(_copyHandler);
        }

        // XAML. Обработка открытых документов
        private void BTN_LoadLinkedDoc_Click(object sender, RoutedEventArgs e)
        {
            additionalDocument = null;

            // Поиск документа
            RevitApp revitApp = _uiapp.Application;
            DocumentSet docSet = revitApp.Documents;
            List<Document> allOpenDocuments = new List<Document>();

            foreach (Document doc in docSet)
            {
                if (!doc.IsFamilyDocument && doc.Title != mainDocument.Title && !doc.IsLinked)
                {
                    allOpenDocuments.Add(doc);
                }
            }

            // Отрисовка интерфейса
            if (allOpenDocuments.Count > 0)
            {
                var selectionWindow = new DocumentSelectionWindow(allOpenDocuments)
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                };

                if (selectionWindow.ShowDialog() == true)
                {
                    SP_OtherFile.Children.Clear();
                    SP_OtherFileSetings.Children.Clear();

                    additionalDocument = selectionWindow.SelectedDocument;

                    if (additionalDocument != null)
                    {
                        string additionalDocName = additionalDocument.Title;

                        StackPanel panel = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            Margin = new Thickness(10, 4, 0, 4)
                        };

                        TextBlock textBlock = new TextBlock
                        {
                            Text = "Имя открытого документа",
                            Height = 15,
                            FontSize = 11
                        };

                        System.Windows.Controls.TextBox textBox = new System.Windows.Controls.TextBox
                        {
                            Name = "TB_OtherDocumentName",
                            Height = 25,
                            Width = 655,
                            HorizontalAlignment = HorizontalAlignment.Left,
                            Margin = new Thickness(0, 0, 15, 0),
                            Padding = new Thickness(3),
                            Cursor = System.Windows.Input.Cursors.Arrow,
                            IsReadOnly = true,
                            Text = additionalDocName,
                            ToolTip = additionalDocument.PathName
                        };

                        panel.Children.Add(textBlock);
                        panel.Children.Add(textBox);
                        SP_OtherFile.Children.Add(panel);

                        // Обработка документа
                        if (mainDocument != null && additionalDocument != null)
                        {
                            BTN_OpenTempalte.IsEnabled = true;

                            DisplayAllTemplatesInPanel(mainDocument, additionalDocument);

                            viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, View>>();
                        }
                        else
                        {
                            MessageBox.Show("Ошибка при обработке документа", "Предупреждение");

                            BTN_OpenTempalte.IsEnabled = false;
                            CHK_ReplaceTypes.IsEnabled = false;
                            CHK_ReplaceTypes.IsChecked = false;
                            CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                            BTN_RunOneFile.IsEnabled = false;
                            BTN_RunManyFile.IsEnabled = false;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Выбор документа отменён.", "Информация");

                    additionalDocument = null;

                    SP_OtherFile.Children.Clear();
                    SP_OtherFileSetings.Children.Clear();
                    BTN_OpenTempalte.IsEnabled = false;
                    CHK_ReplaceTypes.IsEnabled = false;
                    CHK_ReplaceTypes.IsChecked = false;
                    CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                    BTN_RunOneFile.IsEnabled = false;
                    BTN_RunManyFile.IsEnabled = false;
                }
            }
            else
            {
                MessageBox.Show("Нет доступных открытых документов.", "Предупреждение");

                BTN_OpenTempalte.IsEnabled = false;
                CHK_ReplaceTypes.IsEnabled = false;
                CHK_ReplaceTypes.IsChecked = false;
                CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                BTN_RunOneFile.IsEnabled = false;
                BTN_RunManyFile.IsEnabled = false;
            }
        }

        // XAML. Обработка внешних документов
        private void BTN_LoadCustomDoc_Click(object sender, RoutedEventArgs e)
        {
            // Открытие документа
            additionalDocument = null;

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Revit Files (*.rvt)|*.rvt",
                Title = "Выберите проект Revit"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                if (File.Exists(filePath))
                {
                    try
                    {
                        SP_OtherFile.Children.Clear();
                        SP_OtherFileSetings.Children.Clear();

                        _uiapp.DialogBoxShowing += OnDialogBoxShowing;
                        _uiapp.Application.FailuresProcessing += OnFailureProcessing;

                        Document externalDoc = _uiapp.OpenAndActivateDocument(filePath).Document;

                        if (externalDoc != null)
                        {
                            _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                            _uiapp.Application.FailuresProcessing -= OnFailureProcessing;

                            // Отрисовка интерфейса
                            additionalDocument = externalDoc;
                            string additionalDocName = externalDoc.Title;

                            StackPanel panel = new StackPanel
                            {
                                Orientation = Orientation.Vertical,
                                Margin = new Thickness(10, 4, 0, 4)
                            };

                            TextBlock textBlock = new TextBlock
                            {
                                Text = "Имя открытого документа",
                                Height = 15,
                                FontSize = 11
                            };

                            System.Windows.Controls.TextBox textBox = new System.Windows.Controls.TextBox
                            {
                                Name = "TB_OtherDocumentName",
                                Height = 25,
                                Width = 655,
                                HorizontalAlignment = HorizontalAlignment.Left,
                                Margin = new Thickness(0, 0, 15, 0),
                                Padding = new Thickness(3),
                                Cursor = System.Windows.Input.Cursors.Arrow,
                                IsReadOnly = true,
                                Text = additionalDocName,
                                ToolTip = filePath
                            };

                            panel.Children.Add(textBlock);
                            panel.Children.Add(textBox);
                            SP_OtherFile.Children.Add(panel);

                            // Обработка документа
                            if (mainDocument != null && additionalDocument != null)
                            {
                                BTN_OpenTempalte.IsEnabled = true;

                                DisplayAllTemplatesInPanel(mainDocument, additionalDocument);

                                viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, View>>();
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при обработке документа", "Предупреждение");

                                BTN_OpenTempalte.IsEnabled = false;
                                CHK_ReplaceTypes.IsEnabled = false;
                                CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                                BTN_RunOneFile.IsEnabled = false;
                                BTN_RunManyFile.IsEnabled = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка");

                        BTN_OpenTempalte.IsEnabled = false;
                        CHK_ReplaceTypes.IsEnabled = false;
                        CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                        BTN_RunOneFile.IsEnabled = false;
                        BTN_RunManyFile.IsEnabled = false;
                        _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    }
                }
                else
                {
                    MessageBox.Show("Файл не найден.", "Ошибка");

                    BTN_OpenTempalte.IsEnabled = false;
                    CHK_ReplaceTypes.IsEnabled = false;
                    CHK_ReplaceTypes.IsChecked = false;
                    CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                    BTN_RunOneFile.IsEnabled = false;
                    BTN_RunManyFile.IsEnabled = false;

                }
            }
            else
            {
                additionalDocument = null;
                SP_OtherFile.Children.Clear();
                SP_OtherFileSetings.Children.Clear();

                BTN_OpenTempalte.IsEnabled = false;
                CHK_ReplaceTypes.IsEnabled = false;
                CHK_ReplaceTypes.IsChecked = false;
                CHK_ReplaceTypes.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FF7F7F7F"));
                BTN_RunOneFile.IsEnabled = false;
                BTN_RunManyFile.IsEnabled = false;
            }
        }

        // XAML. Кнопка "Обновить шаблоны"
        private void BTN_OpenTempalte_Click(object sender, RoutedEventArgs e)
        {           
            DisplayAllTemplatesInPanel(mainDocument, additionalDocument);
        }

        // Отрисовка интефейса. Шаблоны
        private void DisplayAllTemplatesInPanel(Document mainDocument, Document additionalDocument)
        {
            viewOnlyTemplateChanges.Clear();
            SP_OtherFileSetings.Children.Clear();

            var viewTemplateFilter = new ElementClassFilter(typeof(View));
            var mainCollector = new FilteredElementCollector(mainDocument).WherePasses(viewTemplateFilter).Cast<View>();

            mainTemplates = new SortedDictionary<string, View>();

            foreach (var view in mainCollector)
            {
                if (view.IsTemplate)
                    mainTemplates[view.Name] = view;
            }

            ScrollViewer scrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });   
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(380) });  
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });

            TextBlock emptyHeader = new TextBlock 
            { 
                Text = "", 
                Margin = new Thickness(5) 
            };
            TextBlock nameHeader = new TextBlock
            {
                Text = "Имя шаблона (в активном документе)",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(2, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            TextBlock replaceInViewsHeader = new TextBlock
            {
                Text = "Замена на видах",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(5, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            TextBlock copyHeader = new TextBlock
            {
                Text = "Резервная копия",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(5, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(emptyHeader, 0); Grid.SetColumn(emptyHeader, 0); grid.Children.Add(emptyHeader);
            Grid.SetRow(nameHeader, 0); Grid.SetColumn(nameHeader, 1); grid.Children.Add(nameHeader);
            Grid.SetRow(replaceInViewsHeader, 0); Grid.SetColumn(replaceInViewsHeader, 2); grid.Children.Add(replaceInViewsHeader);
            Grid.SetRow(copyHeader, 0); Grid.SetColumn(copyHeader, 3); grid.Children.Add(copyHeader);

            int row = 1;

            foreach (var kvp in mainTemplates)
            {
                string templateName = kvp.Key;
                View templateView = kvp.Value;

                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                CheckBox selectCheckBox = new CheckBox
                {                   
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    ToolTip = "Выбор шаблона в активном документе"
                };

                TextBlock nameBlock = new TextBlock
                {
                    Text = templateName,
                    FontSize = 11,
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = templateName,
                    Foreground = Brushes.Black
                };

                CheckBox replaceCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    ToolTip = "Заменит данный шаблон вида на всех видах, где он встречается в открытом документе",
                    IsEnabled = false
                };

                CheckBox copyCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    ToolTip = "Если шаблон с таким именем уже существует, то создасться резервная копия старого шаблона с именем 'староеИмяШаблона_ДатаВремя'",
                    IsEnabled = false
                };

                selectCheckBox.Checked += (s, e) =>
                {
                    copyCheckBox.IsEnabled = true;
                    copyCheckBox.IsChecked = true;
                    replaceCheckBox.IsEnabled = true;
                    replaceCheckBox.IsChecked = true;

                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                selectCheckBox.Unchecked += (s, e) =>
                {
                    replaceCheckBox.IsEnabled = false;
                    replaceCheckBox.IsChecked = false;
                    copyCheckBox.IsEnabled = false;
                    copyCheckBox.IsChecked = false;
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                replaceCheckBox.Checked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                replaceCheckBox.Unchecked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                copyCheckBox.Checked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                copyCheckBox.Unchecked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
                };

                Grid.SetRow(selectCheckBox, row);
                Grid.SetColumn(selectCheckBox, 0);
                grid.Children.Add(selectCheckBox);

                Grid.SetRow(nameBlock, row);
                Grid.SetColumn(nameBlock, 1);
                grid.Children.Add(nameBlock);

                Grid.SetRow(replaceCheckBox, row);
                Grid.SetColumn(replaceCheckBox, 2);
                grid.Children.Add(replaceCheckBox);

                Grid.SetRow(copyCheckBox, row);
                Grid.SetColumn(copyCheckBox, 3);
                grid.Children.Add(copyCheckBox);

                row++;

                UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, templateView);
            }

            scrollViewer.Content = grid;
            SP_OtherFileSetings.Children.Add(scrollViewer);

            CHK_ReplaceTypes.IsEnabled = true;
            CHK_ReplaceTypes.IsChecked = true;
            CHK_ReplaceTypes.Foreground = new SolidColorBrush(Colors.Black);

            BTN_RunOneFile.IsEnabled = true;
            BTN_RunManyFile.IsEnabled = true;
        }

        // Виды. Метод записи изменений для шаблонов
        private void UpdateTemplateBackupChoice(string templateName, CheckBox selectCheckBox, CheckBox replaceCheckBox, CheckBox copyCheckBox, View templateView)
        {
            string statusView = selectCheckBox.IsChecked == true ? "resave" : "ignore";
            string statusResaveInView = replaceCheckBox.IsChecked == true ? "resaveIV" : "ignoreIV";
            string statusCopy = copyCheckBox.IsChecked == true ? "copyView" : "ignoreCopyView";

            viewOnlyTemplateChanges[templateName] = new Tuple<string, string, string, View>(statusView, statusResaveInView, statusCopy, templateView);
        }

        // XAML. Кнопка запускающая работу плагина для одного документа
        private void BTN_RunOneFile_Click(object sender, RoutedEventArgs e)
        {

            if (additionalDocument == null || !additionalDocument.IsValidObject)
            {
                MessageBox.Show($"Документ, в который будут копироваться шаблоны видов не выбран или закрыт. Откройте документ, в который вы хотите копировать шаблоны видов, и повторите попытку.", "KPLN. Информация");

                additionalDocument = null;
                SP_OtherFile.Children.Clear();
                SP_OtherFileSetings.Children.Clear();

                return;
            }

            if (viewOnlyTemplateChanges.Count == 0)
            {
                MessageBox.Show($"Нет изменений для приминения", "KPLN. Информация");
                return;
            }

            this.Close();

            _copyHandler.ExecuteAction = (uiApp) =>
            {
                _uiapp.DialogBoxShowing += OnDialogBoxShowing;
                _uiapp.Application.FailuresProcessing += OnFailureProcessing;

                string pathAD = additionalDocument.PathName;
                if (!string.IsNullOrEmpty(pathAD))
                {
                    uiApp.OpenAndActivateDocument(pathAD);
                }

                List<string> errorMessages = new List<string>();

                using (Transaction trans = new Transaction(additionalDocument, "KPLN. Копирование шаблонов видов"))
                {
                    trans.Start();

                    // Шаблоны видов
                    if (viewOnlyTemplateChanges.Count > 0)
                    {
                        foreach (var kvp in viewOnlyTemplateChanges)
                        {
                            string viewTemplateName = kvp.Key;
                            string statusView = kvp.Value.Item1;
                            string statusResaveInView = kvp.Value.Item2;
                            string statusCopyView = kvp.Value.Item3;
                            View templateView = kvp.Value.Item4;

                            if (statusView == "resave" && templateView != null)
                            {
                                View existingTemplate = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

                                List<View> viewsUsingexistingTemplate = null;

                                // Удаление
                                if (statusCopyView == "ignoreCopyView")
                                {
                                    if (CHK_ReplaceTypes.IsChecked == true)
                                    {
                                        RemoveDuplicateTypes(mainDocument, additionalDocument, existingTemplate);
                                    }

                                    if (existingTemplate != null)
                                    {
                                        viewsUsingexistingTemplate = new FilteredElementCollector(additionalDocument)
                                            .OfClass(typeof(View))
                                            .Cast<View>()
                                            .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                            .ToList();

                                        try
                                        {
                                            existingTemplate.Name = $"{existingTemplate.Name}_DeleteTemp{DateTime.Now:yyyyMMddHHmmss}";
                                            additionalDocument.Delete(existingTemplate.Id);
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось удалить шаблон {existingTemplate.ToString()}: {ex.Message}\n");
                                        }
                                    }

                                    CopyPasteOptions options = new CopyPasteOptions();
                                    options.SetDuplicateTypeNamesHandler(new MyDuplicateTypeNamesHandler());

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                        mainDocument,
                                        new List<ElementId> { templateView.Id },
                                        additionalDocument,
                                        null,
                                        options
                                    );
                                  
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                                    string templateName = templateView.Name;

                                    try
                                    {
                                        copiedTemplateViewNew.Name = viewTemplateName;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не удалось переименовать шаблон {copiedTemplateViewNew.Name}: {ex.Message}\n");
                                    }

                                    if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                    {
                                        foreach (View view in viewsUsingexistingTemplate)
                                        {
                                            try
                                            {
                                                view.ViewTemplateId = copiedTemplateIdNew;
                                            }
                                            catch (Exception ex)
                                            {
                                                errorMessages.Add($"Не удалось назначить новый шаблон на вид '{view.Name}': {ex.Message}");
                                            }
                                        }
                                    }
                                }
                                // Резервная копия
                                else if (statusCopyView == "copyView")
                                {
                                    if (CHK_ReplaceTypes.IsChecked == true)
                                    {
                                        RemoveDuplicateTypes(mainDocument, additionalDocument, existingTemplate);
                                    }

                                    if (existingTemplate != null)
                                    {
                                        viewsUsingexistingTemplate = new FilteredElementCollector(additionalDocument)
                                            .OfClass(typeof(View))
                                            .Cast<View>()
                                            .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                            .ToList();

                                        try
                                        {
                                            existingTemplate.Name = $"{viewTemplateName}_{DateTime.Now:yyyyMMdd_HHmmss}";
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось переименовать шаблон {existingTemplate.Name}: {ex.Message}\n");
                                        }
                                    }

                                    CopyPasteOptions options = new CopyPasteOptions();
                                    options.SetDuplicateTypeNamesHandler(new MyDuplicateTypeNamesHandler());

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(
                                        mainDocument,
                                        new List<ElementId> { templateView.Id },
                                        additionalDocument,
                                        null,
                                        options
                                    );

                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                                    copiedTemplateViewNew.Name = viewTemplateName;

                                    if (statusResaveInView == "resaveIV" && viewsUsingexistingTemplate != null)
                                    {
                                        foreach (View view in viewsUsingexistingTemplate)
                                        {
                                            try
                                            {
                                                view.ViewTemplateId = copiedTemplateIdNew;
                                            }
                                            catch (Exception ex)
                                            {
                                                errorMessages.Add($"Не удалось назначить новый шаблон на вид '{view.Name}': {ex.Message}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                 
                    _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    _uiapp.Application.FailuresProcessing -= OnFailureProcessing;

                    trans.Commit();

                    if (errorMessages.Count > 0)
                    {
                        string combined = string.Join("\n", errorMessages);
                        MessageBox.Show($"Применение завершено с ошибками:\n\n{combined}", "KPLN. Завершено с ошибками");
                    }
                    else
                    {
                        MessageBox.Show("Изменения успешно применены к документу.", "KPLN. Готово");
                    }

                    heritageMainDocument = mainDocument;
                    heritageAdditionalDocument = additionalDocument;
;
                    string pathMD = mainDocument.PathName;
                    if (!string.IsNullOrEmpty(pathMD))
                    {
                        uiApp.OpenAndActivateDocument(pathMD);
                    }

                    CopyViewForm copyViewForm = new CopyViewForm(_uiapp, heritageMainDocument, heritageAdditionalDocument);
                    copyViewForm.ShowDialog();
                }           
            };

            _copyEvent.Raise();         
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

        // XAML. Выбор и обработка нескольких документов
        private void BTN_RunManyFile_Click(object sender, RoutedEventArgs e)
        {
            bool replaceTypes = CHK_ReplaceTypes.IsChecked == true;
            var openManeDocsWindows = new ManyDocumentsSelectionWindow(_uiapp, mainDocument, additionalDocument, viewOnlyTemplateChanges, replaceTypes);
            openManeDocsWindows.Owner = this;           
            openManeDocsWindows.ShowDialog();
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
