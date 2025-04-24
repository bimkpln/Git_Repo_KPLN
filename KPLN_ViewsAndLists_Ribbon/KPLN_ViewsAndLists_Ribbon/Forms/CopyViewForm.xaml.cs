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
using ComboBox = System.Windows.Controls.ComboBox;



namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public class CopyViewExternalHandler : IExternalEventHandler
    {
        public Action<UIApplication> ExecuteAction;

        public void Execute(UIApplication app)
        {
            ExecuteAction?.Invoke(app);
        }

        public string GetName() => "KPLN. Copy Views External Handler";
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

        SortedDictionary<string,View> mainDocumentViewAndTemplates;
        SortedDictionary<string, View> additionalDocumentViewAndTemplates;
        SortedDictionary<string, View> additionalDocumentViewsAndTemplatesNT;

        IEnumerable<View> mainDocumentCollectorVAT;
        IEnumerable<View> additionalDocumentCollectorVAT;

        private Dictionary<string, Tuple<string, string, View>> viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
        private Dictionary<string, Tuple<string,string,View>> viewTemplateChanges = new Dictionary<string, Tuple<string,string,View>>();

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
                        BTN_OpenTemplateAndView.IsEnabled = true;

                        DisplayAllTemplatesInPanel(mainDocument, additionalDocument);

                        viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                        viewTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
                        BTN_OpenTempalte.IsEnabled = false;
                        BTN_OpenTemplateAndView.IsEnabled = false;
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
                            BTN_OpenTemplateAndView.IsEnabled = true;

                            CB_ReplaceInViews.IsEnabled = true;
                            CB_ReplaceInViews.Foreground = Brushes.Black;

                            DisplayAllTemplatesInPanel(mainDocument, additionalDocument);

                            viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                            viewTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                        }
                        else
                        {
                            MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
                            CB_ReplaceInViews.IsEnabled = false;
                            CB_ReplaceInViews.IsChecked = false;
                            CB_ReplaceInViews.Foreground = Brushes.Gray;
                            BTN_OpenTempalte.IsEnabled = false;
                            BTN_OpenTemplateAndView.IsEnabled = false;
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

                    CB_ReplaceInViews.IsEnabled = false;
                    CB_ReplaceInViews.IsChecked = false;
                    CB_ReplaceInViews.Foreground = Brushes.Gray;
                    BTN_OpenTempalte.IsEnabled = false;
                    BTN_OpenTemplateAndView.IsEnabled = false;
                    BTN_RunOneFile.IsEnabled = false;
                    BTN_RunManyFile.IsEnabled = false;
                }
            }
            else
            {
                MessageBox.Show("Нет доступных открытых документов.", "Предупреждение");

                CB_ReplaceInViews.IsEnabled = false;
                CB_ReplaceInViews.IsChecked = false;
                CB_ReplaceInViews.Foreground = Brushes.Gray;
                BTN_OpenTempalte.IsEnabled = false;
                BTN_OpenTemplateAndView.IsEnabled = false;
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
                                BTN_OpenTemplateAndView.IsEnabled = true;

                                DisplayAllTemplatesInPanel(mainDocument, additionalDocument);

                                viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                                viewTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
                                CB_ReplaceInViews.IsEnabled = false;
                                CB_ReplaceInViews.IsChecked = false;
                                CB_ReplaceInViews.Foreground = Brushes.Gray;
                                BTN_OpenTempalte.IsEnabled = false;
                                BTN_OpenTemplateAndView.IsEnabled = false;
                                BTN_RunOneFile.IsEnabled = false;
                                BTN_RunManyFile.IsEnabled = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка");
                        CB_ReplaceInViews.IsEnabled = false;
                        CB_ReplaceInViews.IsChecked = false;
                        CB_ReplaceInViews.Foreground = Brushes.Gray;
                        BTN_OpenTempalte.IsEnabled = false;
                        BTN_OpenTemplateAndView.IsEnabled = false;
                        BTN_RunOneFile.IsEnabled = false;
                        BTN_RunManyFile.IsEnabled = false;
                        _uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    }
                }
                else
                {
                    MessageBox.Show("Файл не найден.", "Ошибка");

                    CB_ReplaceInViews.IsEnabled = false;
                    CB_ReplaceInViews.IsChecked = false;
                    CB_ReplaceInViews.Foreground = Brushes.Gray;
                    BTN_OpenTempalte.IsEnabled = false;
                    BTN_OpenTemplateAndView.IsEnabled = false;
                    BTN_RunOneFile.IsEnabled = false;
                    BTN_RunManyFile.IsEnabled = false;

                }
            }
            else
            {
                additionalDocument = null;
                SP_OtherFile.Children.Clear();
                SP_OtherFileSetings.Children.Clear();

                CB_ReplaceInViews.IsEnabled = false;
                CB_ReplaceInViews.IsChecked = false;
                CB_ReplaceInViews.Foreground = Brushes.Gray;
                BTN_OpenTempalte.IsEnabled = false;
                BTN_OpenTemplateAndView.IsEnabled = false;
                BTN_RunOneFile.IsEnabled = false;
                BTN_RunManyFile.IsEnabled = false;
            }
        }

        // XAML. Кнопка "Шаблоны"
        private void BTN_OpenTempalte_Click(object sender, RoutedEventArgs e)
        {           
            DisplayAllTemplatesInPanel(mainDocument, additionalDocument);
        }

        // XAML. Кнопка "Шаблоны и виды"
        private void BTN_OpenTemplateAndView_Click(object sender, RoutedEventArgs e)
        {          
            DisplayAllViewsAndTemplatesInPanel(mainDocument, additionalDocument);
        }

        // Отрисовка интефейса. Шаблоны
        private void DisplayAllTemplatesInPanel(Document mainDocument, Document additionalDocument)
        {
            CB_ReplaceInViews.IsEnabled = true;
            CB_ReplaceInViews.Foreground = Brushes.Black;

            viewTemplateChanges.Clear();
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
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(500) });  
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });

            TextBlock emptyHeader = new TextBlock 
            { 
                Text = "", 
                Margin = new Thickness(5) 
            };
            TextBlock nameHeader = new TextBlock
            {
                Text = "Имя шаблона",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(2, 3, 0, 3),
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
            Grid.SetRow(emptyHeader, 0);
            Grid.SetColumn(emptyHeader, 0);
            grid.Children.Add(emptyHeader);
            Grid.SetRow(nameHeader, 0);
            Grid.SetColumn(nameHeader, 1);
            grid.Children.Add(nameHeader);
            Grid.SetRow(copyHeader, 0);
            Grid.SetColumn(copyHeader, 2);
            grid.Children.Add(copyHeader);

            int row = 1;

            foreach (var kvp in mainTemplates)
            {
                string templateName = kvp.Key;
                View templateView = kvp.Value;

                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                CheckBox selectCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5)
                };

                bool isAlsoInAdditional = new FilteredElementCollector(additionalDocument)
                    .OfClass(typeof(View)).Cast<View>().Any(v => v.IsTemplate && v.Name == templateName);

                TextBlock nameBlock = new TextBlock
                {
                    Text = templateName,
                    FontSize = 11,
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = templateName,
                    Foreground = isAlsoInAdditional ? Brushes.Blue : Brushes.Black
                };

                CheckBox copyCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    IsEnabled = false
                };

                selectCheckBox.Checked += (s, e) =>
                {
                    copyCheckBox.IsEnabled = true;
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, copyCheckBox, templateView);
                };

                selectCheckBox.Unchecked += (s, e) =>
                {
                    copyCheckBox.IsEnabled = false;
                    copyCheckBox.IsChecked = false;
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, copyCheckBox, templateView);
                };

                copyCheckBox.Checked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, copyCheckBox, templateView);
                };

                copyCheckBox.Unchecked += (s, e) =>
                {
                    UpdateTemplateBackupChoice(templateName, selectCheckBox, copyCheckBox, templateView);
                };

                Grid.SetRow(selectCheckBox, row);
                Grid.SetColumn(selectCheckBox, 0);
                grid.Children.Add(selectCheckBox);

                Grid.SetRow(nameBlock, row);
                Grid.SetColumn(nameBlock, 1);
                grid.Children.Add(nameBlock);

                Grid.SetRow(copyCheckBox, row);
                Grid.SetColumn(copyCheckBox, 2);
                grid.Children.Add(copyCheckBox);

                row++;

                UpdateTemplateBackupChoice(templateName, selectCheckBox, copyCheckBox, templateView);
            }

            scrollViewer.Content = grid;
            SP_OtherFileSetings.Children.Add(scrollViewer);
            BTN_RunOneFile.IsEnabled = true;
            BTN_RunManyFile.IsEnabled = true;
        }

        // Виды. Метод записи изменений для шаблонов
        private void UpdateTemplateBackupChoice(string templateName, CheckBox statusCheckBox, CheckBox copyCheckBox, View templateView)
        {
            string statusView = statusCheckBox.IsChecked == true ? "resave" : "ignore";
            string statusCopy = copyCheckBox.IsChecked == true ? "copyView" : "ignoreCopyView";

            viewOnlyTemplateChanges[templateName] = new Tuple<string, string, View>(statusView, statusCopy, templateView);
        }

        // Отрисовка интефейса. Виды и шаблоны
        private void DisplayAllViewsAndTemplatesInPanel(Document mainDocument, Document additionalDocument)
        {
            CB_ReplaceInViews.IsEnabled = false;
            CB_ReplaceInViews.IsChecked = false;
            CB_ReplaceInViews.Foreground = Brushes.Gray;

            viewTemplateChanges.Clear();
            viewOnlyTemplateChanges.Clear();
            SP_OtherFileSetings.Children.Clear();
         
            var viewTemplateFilter = new ElementClassFilter(typeof(View));

            mainDocumentViewAndTemplates = new SortedDictionary<string, View>();
            mainDocumentCollectorVAT = new FilteredElementCollector(mainDocument).WherePasses(viewTemplateFilter).Cast<View>();

            foreach (var view in mainDocumentCollectorVAT)
            {
                if (view.IsTemplate)
                {
                    mainDocumentViewAndTemplates[view.Name] = view;
                }          
            }

            additionalDocumentViewAndTemplates = new SortedDictionary<string, View>();
            additionalDocumentViewsAndTemplatesNT = new SortedDictionary<string, View>();
            additionalDocumentCollectorVAT = new FilteredElementCollector(additionalDocument).WherePasses(viewTemplateFilter).Cast<View>();

            foreach (var view in additionalDocumentCollectorVAT)
            {
                if (!view.IsTemplate)
                {
                    additionalDocumentViewsAndTemplatesNT[view.Name] = view;
                }
                else
                {
                    additionalDocumentViewAndTemplates[view.Name] = view;
                }
            }

            ScrollViewer scrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(285) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(68) });

            TextBlock nameHeader = new TextBlock
            {
                Text = "Имя вида",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(2, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            TextBlock templateHeader = new TextBlock
            {
                Text = "Шаблон",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(5, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            TextBlock copyTemplateHeader = new TextBlock
            {
                Text = "Резерв. копия",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 3, 0, 3),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(nameHeader, 0);
            Grid.SetColumn(nameHeader, 0);
            grid.Children.Add(nameHeader);
            Grid.SetRow(templateHeader, 0);
            Grid.SetColumn(templateHeader, 1);
            grid.Children.Add(templateHeader);
            Grid.SetRow(copyTemplateHeader, 0);
            Grid.SetColumn(copyTemplateHeader, 2);
            grid.Children.Add(copyTemplateHeader);
            int row = 1;

            foreach (var kvp in additionalDocumentViewsAndTemplatesNT)
            {
                View view = kvp.Value;

                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                TextBlock viewNameBlock = new TextBlock
                {
                    Text = view.Name,
                    FontSize = 11,
                    Margin = new Thickness(5),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = view.Name
                };

                CheckBox templateCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5, 0, 5, 0),
                    IsChecked = false
                };
             
                string initialTemplateName = "Нет шаблона";
                bool isFromMain = false;
                bool isFromAdditional = false;

                if (view.ViewTemplateId != ElementId.InvalidElementId)
                {
                    var templateView = additionalDocument.GetElement(view.ViewTemplateId) as View;
                    if (templateView != null)
                    {
                        initialTemplateName = templateView.Name;
                        isFromMain = mainDocumentViewAndTemplates.ContainsKey(initialTemplateName);
                        isFromAdditional = additionalDocumentViewAndTemplates.ContainsKey(initialTemplateName);
                    }
                }

                List<string> templateList = new List<string> { "Нет шаблона" };
                templateList.AddRange(mainDocumentViewAndTemplates.Keys);

                if (!mainDocumentViewAndTemplates.ContainsKey(initialTemplateName) &&
                    additionalDocumentViewAndTemplates.ContainsKey(initialTemplateName) &&
                    !templateList.Contains(initialTemplateName))
                {
                    templateList.Add(initialTemplateName); 
                }

                bool isTemplateUnknown = !isFromMain && isFromAdditional;

                ComboBox templateComboBox = new ComboBox
                {
                    FontSize = 11,
                    Margin = new Thickness(0, 2, 5, 2),
                    Width = 250,
                    ItemsSource = templateList,
                    IsEnabled = false,
                    IsEditable = true,
                    Foreground = isTemplateUnknown ? Brushes.Red : Brushes.Black
                };

                if (isFromMain)
                {
                    templateComboBox.SelectedItem = initialTemplateName;
                }
                else if (isFromAdditional)
                {
                    templateComboBox.Text = initialTemplateName;
                    templateComboBox.SelectedItem = null;
                }
                else
                {
                    templateComboBox.SelectedItem = "Нет шаблона";
                }

                templateComboBox.Text = initialTemplateName;
                string lockedTemplate = initialTemplateName;

                CheckBox copyTemplateCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Margin = new Thickness(0, 5, 5, 5),
                    IsEnabled = false,
                };

                templateCheckBox.Checked += (s, e) =>
                {
                    templateComboBox.IsEnabled = true;
                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
                    UpdateCopyCheckBox(templateCheckBox, templateComboBox, copyTemplateCheckBox);
                };

                templateCheckBox.Unchecked += (s, e) =>
                {
                    templateComboBox.IsEnabled = false;
                    copyTemplateCheckBox.IsEnabled = false;

                    if (mainDocumentViewAndTemplates.ContainsKey(lockedTemplate))
                    {
                        templateComboBox.SelectedItem = lockedTemplate;
                        templateComboBox.Foreground = Brushes.Black;
                    }
                    else if (additionalDocumentViewAndTemplates.ContainsKey(lockedTemplate))
                    {
                        templateComboBox.SelectedItem = lockedTemplate;
                        templateComboBox.Foreground = Brushes.Red;
                    }
                    else
                    {
                        templateComboBox.SelectedItem = "Нет шаблона";
                        templateComboBox.Foreground = Brushes.Black;
                    }

                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
                };

                templateComboBox.SelectionChanged += (s, e) =>
                {
                    if (templateComboBox.SelectedItem is string selected)
                    {
                        if (selected == "Не выбрано")
                        {
                            templateComboBox.Foreground = Brushes.Black;
                        }
                        else if (mainDocumentViewAndTemplates.ContainsKey(selected))
                        {
                            templateComboBox.Foreground = Brushes.Black;
                        }
                        else
                        {
                            templateComboBox.SelectedItem = "Не выбрано";
                            templateComboBox.Foreground = Brushes.Black;
                        }
                    }

                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
                    UpdateCopyCheckBox(templateCheckBox, templateComboBox, copyTemplateCheckBox);
                };

                copyTemplateCheckBox.Checked += (s, e) =>
                {
                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
                    UpdateCopyCheckBox(templateCheckBox, templateComboBox, copyTemplateCheckBox);
                };

                copyTemplateCheckBox.Unchecked += (s, e) =>
                {
                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
                    UpdateCopyCheckBox(templateCheckBox, templateComboBox, copyTemplateCheckBox);
                };

                StackPanel templatePanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    VerticalAlignment = VerticalAlignment.Center
                };

                templatePanel.Children.Add(templateCheckBox);
                templatePanel.Children.Add(templateComboBox);
                Grid.SetRow(viewNameBlock, row);
                Grid.SetColumn(viewNameBlock, 0);
                grid.Children.Add(viewNameBlock);
                Grid.SetRow(templatePanel, row);
                Grid.SetColumn(templatePanel, 1);
                grid.Children.Add(templatePanel);
                Grid.SetRow(copyTemplateCheckBox, row);
                Grid.SetColumn(copyTemplateCheckBox, 2);
                grid.Children.Add(copyTemplateCheckBox);

                row++;

                UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox, copyTemplateCheckBox);
            }

            scrollViewer.Content = grid;
            SP_OtherFileSetings.Children.Add(scrollViewer);
            BTN_RunOneFile.IsEnabled = true;
            BTN_RunManyFile.IsEnabled = true;
        }

        // Виды и шаблоны. Вспомогательный метод проверки данных CheckBox
        void UpdateCopyCheckBox(CheckBox templateCheckBox, ComboBox templateComboBox, CheckBox copyTemplateCheckBox)
        {
            string selectedName = templateComboBox.SelectedItem as string ?? templateComboBox.Text;
            if (templateCheckBox.IsChecked == true && selectedName != "Нет шаблона")
            {
                copyTemplateCheckBox.IsEnabled = true;
            }
            else
            {
                copyTemplateCheckBox.IsEnabled = false;
            }
        }

        // Виды и шаблоны. Вспомогательный метод записи данных с интерфейса в словарь
        void UpdateTemplateChange(string viewName, ComboBox comboBox, CheckBox statusCheckBox, CheckBox statusCopyCheckBox)
        {
            string selectedName = comboBox.SelectedItem as string ?? comboBox.Text;

            View selectedTemplate = null;
            mainDocumentViewAndTemplates.TryGetValue(selectedName, out selectedTemplate);

            string statusView;            
            if (statusCheckBox.IsChecked == true)
            {
                if (selectedName == "Нет шаблона")
                {
                    statusView = "delete";
                }
                else
                {
                    statusView = "resave";                    
                }
            }
            else
            {
                if (viewTemplateChanges.ContainsKey(viewName))
                {
                    statusView = viewTemplateChanges[viewName].Item1;
                }
                else
                {
                    statusView = "ignore";
                }
            }

            string statusCopyView;
            if (statusCopyCheckBox.IsChecked == true)
            {
                statusCopyView = "copyView";
            }
            else
            {
                if (viewTemplateChanges.ContainsKey(viewName))
                {
                    statusCopyView = viewTemplateChanges[viewName].Item2;
                }
                else
                {
                    statusCopyView = "ignoreCopyView";
                }
            }

            viewTemplateChanges[viewName] = new Tuple<string, string, View>(statusView, statusCopyView, selectedTemplate);
        }










        // XAML. Кнопка запускающая работу плагина. Один файл
        private void BTN_RunOneFile_Click(object sender, RoutedEventArgs e)
        {         
            if (viewOnlyTemplateChanges.Count == 0 && viewTemplateChanges.Count == 0)
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
                            string statusCopyView = kvp.Value.Item2;
                            View templateView = kvp.Value.Item3;

                            if (statusView == "resave" && templateView != null)
                            {
                                // Удаление
                                if (statusCopyView == "ignoreCopyView")
                                {
                                    var existingTemplate = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

                                    List<View> viewsUsingexistingTemplate = null;

                                    if (existingTemplate != null)
                                    {
                                        viewsUsingexistingTemplate = new FilteredElementCollector(additionalDocument)
                                            .OfClass(typeof(View))
                                            .Cast<View>()
                                            .Where(v => !v.IsTemplate && v.ViewTemplateId == existingTemplate.Id)
                                            .ToList();

                                        try
                                        {
                                            existingTemplate.Name = $"{existingTemplate.Name}_temp{DateTime.Now:yyyyMMddHHmmss}";
                                            additionalDocument.Delete(existingTemplate.Id);
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось удалить шаблон {existingTemplate.ToString()}: {ex.Message}\n");
                                        }
                                    }

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                                    string templateName = templateView.Name;

                                    try
                                    {
                                        copiedTemplateViewNew.Name = viewTemplateName;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не переименовать шаблон {copiedTemplateViewNew.Name}: {ex.Message}\n");
                                    }

                                    if (CB_ReplaceInViews.IsChecked == true && viewsUsingexistingTemplate != null)
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
                                    var existingTemplate = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

                                    List<View> viewsUsingexistingTemplate = null;

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

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;

                                    copiedTemplateViewNew.Name = viewTemplateName;

                                    if (CB_ReplaceInViews.IsChecked == true && viewsUsingexistingTemplate != null)
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








                    // Шаблоны на видах
                    if (viewTemplateChanges.Count > 0)
                    {
                        foreach (var kvp in viewTemplateChanges)
                        {
                            string viewName = kvp.Key;
                            string statusView = kvp.Value.Item1;
                            string statusCopyView = kvp.Value.Item2;
                            View templateView = kvp.Value.Item3;

                            View viewToUpdate = new FilteredElementCollector(additionalDocument)
                                .OfClass(typeof(View))
                                .Cast<View>()
                                .FirstOrDefault(v => !v.IsTemplate && v.Name == viewName);

                            if (viewToUpdate == null)
                                continue;

                            // Удаление шаблона на виде
                            if (statusView == "delete")
                            {
                                try
                                {
                                    viewToUpdate.ViewTemplateId = ElementId.InvalidElementId;
                                }
                                catch (Exception ex)
                                {
                                    errorMessages.Add($"Не удалось удалить шаблон на виде {viewToUpdate.Name}: {ex.Message}\n");
                                }
                            }

                            // Перезапись шаблона вида
                            else if (statusView == "resave")
                            {                              
                                if (statusCopyView == "ignoreCopyView")
                                {
                                    List<View> viewsUsingTemplateFD = null;
                                    List<View> viewsUsingTemplateFDD = null;

                                    List<View> viewsUsingSameTemplateName = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View))
                                        .Cast<View>()
                                        .Where(v =>
                                            !v.IsTemplate &&
                                            v.ViewTemplateId != ElementId.InvalidElementId &&
                                            additionalDocument.GetElement(v.ViewTemplateId) is View vt &&
                                            vt.IsTemplate &&
                                            vt.Name == templateView.Name)
                                        .ToList();

                                    List<View> viewsWithSameTemplateAsViewName = new List<View>();

                                    View viewNamed = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View))
                                        .Cast<View>()
                                        .FirstOrDefault(v => !v.IsTemplate && v.Name == viewName);

                                    if (viewNamed != null && viewNamed.ViewTemplateId != ElementId.InvalidElementId)
                                    {
                                        ElementId templateId = viewNamed.ViewTemplateId;
                                        View template = additionalDocument.GetElement(templateId) as View;

                                        if (template != null && template.IsTemplate)
                                        {
                                            viewsWithSameTemplateAsViewName = new FilteredElementCollector(additionalDocument)
                                                .OfClass(typeof(View))
                                                .Cast<View>()
                                                .Where(v =>
                                                    !v.IsTemplate &&
                                                    v.ViewTemplateId == template.Id)
                                                .ToList();
                                        }
                                    }

                                    List<View> allUniqueViews = viewsWithSameTemplateAsViewName
                                        .Concat(viewsUsingSameTemplateName)
                                        .GroupBy(v => v.Id)
                                        .Select(g => g.First())
                                        .ToList();

                                    // Удаление шаблонов из словаря
                                    ElementId templateIdFD = viewToUpdate.ViewTemplateId;
                                    View templateViewFD = additionalDocument.GetElement(templateIdFD) as View;

                                    viewsUsingTemplateFD = new FilteredElementCollector(additionalDocument).OfClass(typeof(View)).Cast<View>()
                                        .Where(v => !v.IsTemplate && v.ViewTemplateId == templateIdFD).ToList();

                                    foreach (var view in viewsUsingTemplateFD)
                                    {
                                        try
                                        {
                                            view.ViewTemplateId = ElementId.InvalidElementId;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось удалить шаблон {view.ViewTemplateId.ToString()} у вида {view}: {ex.Message}\n");
                                        }
                                    }

                                    try
                                    {
                                        if (templateViewFD != null)
                                        {
                                            templateViewFD.Name = $"{templateViewFD.Name}_temp{DateTime.Now:yyyyMMddHHmmss}";
                                            additionalDocument.Delete(templateIdFD);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не удалось удалить шаблон '{templateViewFD.Name}': {ex.Message}");
                                    }

                                    // Удаление шаблонов из файла
                                    View matchingTemplateInDoc = new FilteredElementCollector(additionalDocument)
                                        .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == templateView.Name);

                                    if (matchingTemplateInDoc != null)
                                    {
                                        ElementId matchingTemplateId = matchingTemplateInDoc.Id;

                                        viewsUsingTemplateFDD = new FilteredElementCollector(additionalDocument).OfClass(typeof(View)).Cast<View>()
                                        .Where(v => !v.IsTemplate && v.ViewTemplateId == matchingTemplateId).ToList();

                                        foreach (var view in viewsUsingTemplateFDD)
                                        {
                                            try
                                            {
                                                view.ViewTemplateId = ElementId.InvalidElementId;
                                            }
                                            catch (Exception ex)
                                            {
                                                errorMessages.Add($"Не удалось удалить шаблон {view.ViewTemplateId.ToString()} у вида {view}: {ex.Message}\n");
                                            }
                                        }

                                        try
                                        {
                                            matchingTemplateInDoc.Name = $"{templateView.Name}_temp{DateTime.Now:yyyyMMddHHmmss}";
                                            additionalDocument.Delete(matchingTemplateId);                                        
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось удалить шаблон '{templateView.Name}': {ex.Message}");
                                        }
                                    }

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;

                                    foreach (var v in allUniqueViews)
                                    {
                                        try
                                        {
                                            v.ViewTemplateId = copiedTemplateIdNew;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид {v.Name}: {ex.Message}\n");
                                        }
                                    }
                                    
                                    try
                                    {
                                        copiedTemplateViewNew.Name = templateView.Name;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не удалось переименовать шаблон вида {copiedTemplateViewNew.Name}: {ex.Message}\n");
                                    }

                                    try
                                    {
                                        viewToUpdate.ViewTemplateId = copiedTemplateIdNew;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид {viewToUpdate.Name}: {ex.Message}\n");
                                    }               
                                    
                                    if (viewToUpdate.ViewType != copiedTemplateViewNew.ViewType)
                                    {
                                        errorMessages.Add($"Типы видов не совпадают: {viewToUpdate.ViewType} ≠ {copiedTemplateViewNew.ViewType}.");

                                        try
                                        {
                                            additionalDocument.Delete(copiedTemplateIdNew);
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Ошибка при удалении шаблона '{copiedTemplateViewNew.Name}': {ex.Message}");
                                        }
                                    }
                                }












                                else if (statusCopyView == "copyView")
                                {                                
                                    View templateViewCopy = additionalDocument.GetElement(viewToUpdate.ViewTemplateId) as View;
                                    string newViewTemplateName;

                                    if (templateViewCopy != null)
                                    {
                                        newViewTemplateName = $"{templateViewCopy.Name}_{DateTime.Now:yyyyMMdd_HHmmss}";
                                    }
                                    else
                                    {
                                        newViewTemplateName = "Не выбрано";
                                    }

                                    View existingTemplate = null; 

                                    if (templateViewCopy != null)
                                    {
                                        existingTemplate = new FilteredElementCollector(additionalDocument)
                                            .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == templateViewCopy.Name);
                                    }

                                    if (existingTemplate != null)
                                    {
                                        try
                                        {
                                            existingTemplate.Name = newViewTemplateName;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось переименовать шаблон {existingTemplate.Name}: {ex.Message}\n");
                                        }
                                    }

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;

                                    try
                                    {
                                        viewToUpdate.ViewTemplateId = copiedTemplateIdNew;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид {viewToUpdate.Name}: {ex.Message}\n");
                                    }

                                    List <View> viewsUsingTemplate = null;
                                    foreach (var v in viewsUsingTemplate)
                                    {
                                        try
                                        {
                                            v.ViewTemplateId = copiedTemplateIdNew;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось назначить шаблон {copiedTemplateViewNew.Name} на вид {v.Name}: {ex.Message}\n");
                                        }
                                    }

                                    if (copiedTemplateViewNew != null)
                                    {
                                        try
                                        {
                                            copiedTemplateViewNew.Name = templateView.Name;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorMessages.Add($"Не удалось переименовать шаблон {copiedTemplateViewNew.Name}: {ex.Message}\n");
                                        }
                                    }
                                }



                                else
                                {
                                    errorMessages.Add($"❌ Вид \"{viewToUpdate.Name}\" не совместим с шаблоном \"{templateView.Name}\"");
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














        // XAML. Кнопка запускающая работу плагина. Несколько файлов
        private void BTN_RunManyFile_Click(object sender, RoutedEventArgs e)
        {

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
