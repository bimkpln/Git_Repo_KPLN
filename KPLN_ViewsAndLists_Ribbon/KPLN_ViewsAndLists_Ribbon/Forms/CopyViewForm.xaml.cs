using System;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Interop;
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
using System.Text;


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

        UIApplication uiapp;
        private readonly DBRevitDialog[] _dbRevitDialogs;

        Document mainDocument = null;
        Document additionalDocument = null;

        SortedDictionary<string, View> mainTemplates;

        SortedDictionary<string,View> mainDocumentViewAndTemplates;
        SortedDictionary<string, View> additionalDocumentViewAndTemplates;
        SortedDictionary<string, View> additionalDocumentViewsAndTemplatesNT;

        IEnumerable<View> mainDocumentCollectorVAT;
        IEnumerable<View> additionalDocumentCollectorVAT;

        private Dictionary<string, Tuple<string, string, View>> viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, View>>();
        private Dictionary<string, Tuple<string,string,View>> viewTemplateChanges = new Dictionary<string, Tuple<string,string,View>>();

        public CopyViewForm(UIApplication uiapp)
        {
            InitializeComponent();
            this.uiapp = uiapp;

            mainDocument = uiapp.ActiveUIDocument.Document;
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
            RevitApp revitApp = uiapp.Application;
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
                        this.Hide();
                        this.Show();
                        this.Topmost = true;
                        this.Activate();

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
                            BTN_Run.IsEnabled = false;
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
                    BTN_OpenTemplateAndView.IsEnabled = false;
                    BTN_Run.IsEnabled = false;
                }
            }
            else
            {
                MessageBox.Show("Нет доступных открытых документов.", "Предупреждение");

                BTN_OpenTempalte.IsEnabled = false;
                BTN_OpenTemplateAndView.IsEnabled = false;
                BTN_Run.IsEnabled = false;
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

                        RevitApp revitApp = uiapp.Application;
                        uiapp.DialogBoxShowing += OnDialogBoxShowing;
                        uiapp.Application.FailuresProcessing += OnFailureProcessing;

                        this.Hide();
                        Document externalDoc = uiapp.OpenAndActivateDocument(filePath).Document;

                        if (externalDoc != null)
                        {
                            uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                            uiapp.Application.FailuresProcessing -= OnFailureProcessing;

                            System.Threading.Thread.Sleep(200); 
                            this.Show();
                            this.Topmost = true;
                            this.Activate();

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
                                BTN_OpenTempalte.IsEnabled = false;
                                BTN_OpenTemplateAndView.IsEnabled = false;
                                BTN_Run.IsEnabled = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка");
                        BTN_OpenTempalte.IsEnabled = false;
                        BTN_OpenTemplateAndView.IsEnabled = false;
                        BTN_Run.IsEnabled = false;
                        uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    }
                }
                else
                {
                    MessageBox.Show("Файл не найден.", "Ошибка");
                }
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
            BTN_Run.IsEnabled = true;
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
            BTN_Run.IsEnabled = true;
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

        // XAML. Кнопка запускающая работу плагина
        private void BTN_Run_Click(object sender, RoutedEventArgs e)
        {
            _copyHandler.ExecuteAction = (uiApp) =>
            {
                List<string> errorMessages = new List<string>();

                using (Transaction trans = new Transaction(additionalDocument, "KPLN. Копирование шаблонов видов"))
                {
                    trans.Start();

                    foreach (var kvp in viewOnlyTemplateChanges)
                    {
                        string viewTemplateName = kvp.Key;
                        string statusView = kvp.Value.Item1;
                        string statusCopyView = kvp.Value.Item2;
                        View templateView = kvp.Value.Item3;

                        if (statusView == "resave" && templateView != null)
                        {                          
                            if (statusCopyView == "ignoreCopyView")
                            {
                                var existingTemplate = new FilteredElementCollector(additionalDocument)
                                    .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

                                if (existingTemplate != null)
                                {
                                    try
                                    {
                                        additionalDocument.Delete(existingTemplate.Id);
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"❌ Не удалось удалить шаблон \"{existingTemplate.ToString()}\": {ex.Message}");
                                    }
                                }

                                ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                                string templateName = templateView.Name;

                                copiedTemplateViewNew.Name = viewTemplateName;
                            }
                            else if (statusCopyView == "copyView")
                            {
                                var existingTemplate = new FilteredElementCollector(additionalDocument)
                                    .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == viewTemplateName);

                                if (existingTemplate != null)
                                {
                                    try
                                    {
                                        existingTemplate.Name = $"{viewTemplateName}_{DateTime.Now:yyyyMMdd_HHmmss}";
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Ошибка", $"Не удалось переименовать шаблон \"{existingTemplate.Name}\": {ex.Message}");
                                    }

                                    ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                                    ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                                    View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                                    string templateName = templateView.Name;

                                    copiedTemplateViewNew.Name = viewTemplateName;
                                }
                            }
                        }
                    }

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

                        ElementId oldTemplateViewId = viewToUpdate.ViewTemplateId;

                        if (viewToUpdate == null)
                            continue;

                        if (statusView == "delete")
                        {
                            viewToUpdate.ViewTemplateId = ElementId.InvalidElementId;
                        }

                        else if (statusView == "resave" && templateView != null)
                        {
                            ICollection<ElementId> copiedIds = ElementTransformUtils.CopyElements(mainDocument, new List<ElementId> { templateView.Id }, additionalDocument, null, new CopyPasteOptions());
                            ElementId copiedTemplateIdNew = copiedIds.FirstOrDefault();
                            View copiedTemplateViewNew = additionalDocument.GetElement(copiedTemplateIdNew) as View;
                            string templateName = templateView.Name;

                            var viewsUsingTemplate = new FilteredElementCollector(additionalDocument).OfClass(typeof(View)).Cast<View>()
                                .Where(v => !v.IsTemplate && v.ViewTemplateId != ElementId.InvalidElementId && additionalDocument.GetElement(v.ViewTemplateId) is View vt && vt.Name == templateName).ToList();

                            if (statusCopyView == "ignoreCopyView")
                            {
                                if (oldTemplateViewId != ElementId.InvalidElementId)
                                {
                                    try
                                    {
                                        additionalDocument.Delete(oldTemplateViewId);
                                    }
                                    catch (Exception ex)
                                    {
                                        errorMessages.Add($"❌ Не удалось удалить шаблон \"{oldTemplateViewId.ToString()}\": {ex.Message}");
                                    }
                                }

                                foreach (var v in viewsUsingTemplate)
                                {
                                    v.ViewTemplateId = copiedTemplateIdNew;
                                }

                                copiedTemplateViewNew.Name = templateName;
                                viewToUpdate.ViewTemplateId = copiedTemplateIdNew;                         
                            }
                            else if (statusCopyView == "copyView")
                            {
                                string newViewTemplateName = $"{templateView.Name}_{DateTime.Now:yyyyMMdd_HHmmss}";

                                View existingTemplate = new FilteredElementCollector(additionalDocument)
                                    .OfClass(typeof(View)).Cast<View>().FirstOrDefault(v => v.IsTemplate && v.Name == templateView.Name);

                                if (existingTemplate != null)
                                {
                                    try
                                    {
                                        existingTemplate.Name = newViewTemplateName;
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Ошибка", $"Не удалось переименовать шаблон \"{existingTemplate.Name}\": {ex.Message}");
                                    }
                                }

                                viewToUpdate.ViewTemplateId = copiedTemplateIdNew;

                                foreach (var v in viewsUsingTemplate)
                                {
                                    v.ViewTemplateId = copiedTemplateIdNew;
                                }
                              
                                if (copiedTemplateViewNew != null)
                                {
                                    try
                                    {
                                        copiedTemplateViewNew.Name = templateName;
                                    }
                                    catch (Exception ex)
                                    {
                                        TaskDialog.Show("Ошибка", $"Не удалось переименовать шаблон \"{copiedTemplateViewNew.Name}\": {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                errorMessages.Add($"❌ Вид \"{viewToUpdate.Name}\" не совместим с шаблоном \"{templateView.Name}\"");
                            }
                        }
                    }

                    trans.Commit();
                }

                if (errorMessages.Count > 0)
                {
                    string combined = string.Join("\n", errorMessages);
                    MessageBox.Show($"Применение завершено с ошибками:\n\n{combined}", "KPLN. Завершено с ошибками");
                }
                else
                {
                    MessageBox.Show("Изменения успешно применены к документу.", "KPLN. Готово");
                }
            };
            _copyEvent.Raise();
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
