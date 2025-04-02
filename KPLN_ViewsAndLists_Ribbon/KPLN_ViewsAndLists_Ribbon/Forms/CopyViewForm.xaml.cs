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
    /// <summary>
    /// Логика взаимодействия для CopyViewForm.xaml
    /// </summary>
    public partial class CopyViewForm : Window
    {
        UIApplication uiapp;
        private readonly DBRevitDialog[] _dbRevitDialogs;

        Document mainDocument = null;
        Document additionalDocument = null;

        SortedDictionary<string,View> mainDocumentTemplates;
        SortedDictionary<string, View> additionalDocumentTemplates;
        SortedDictionary<string, View> additionalDocumentViews;

        IEnumerable<View> mainDocumentCollector;
        IEnumerable<View> additionalDocumentCollector;

        private Dictionary<string, Tuple<string, View>> viewTemplateChanges = new Dictionary<string, Tuple<string, View>>();

        public CopyViewForm(UIApplication uiapp)
        {
            InitializeComponent();

            IntPtr revitHandle = uiapp.MainWindowHandle;
            new WindowInteropHelper(this).Owner = revitHandle;

            this.uiapp = uiapp;

            mainDocument = uiapp.ActiveUIDocument.Document;
            string mainDocPath = mainDocument.PathName;
            string mainDocName = mainDocument.Title;

            if (mainDocPath != null && mainDocName != null)
            {
                TB_ActivDocumentName.Text = $"{mainDocName}";
                TB_ActivDocumentName.ToolTip = $"{mainDocPath}";
            }
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
                    Topmost = true
                };

                if (selectionWindow.ShowDialog() == true)
                {
                    SP_OtherFile.Children.Clear();
                    SP_OtherFileSetings.Children.Clear();

                    additionalDocument = selectionWindow.SelectedDocument;

                    if (additionalDocument != null)
                    {
                        MessageBox.Show($"Выбранный документ загружен:\n{additionalDocument.Title}", "Информация");

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
                            DisplayAllViewsInPanel(mainDocument, additionalDocument);
                        }
                        else
                        {
                            MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
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
                    BTN_Run.IsEnabled = false;
                }
            }
            else
            {
                MessageBox.Show("Нет доступных открытых документов.", "Предупреждение");
                BTN_Run.IsEnabled = false;
            }
        }

        // XAML. Обработка внешних документов
        private void BTN_LoadCustomDoc_Click(object sender, RoutedEventArgs e)
        {
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

                        Document externalDoc = uiapp.OpenAndActivateDocument(filePath).Document;
                        
                        if (externalDoc != null)
                        {
                            MessageBox.Show($"Документ открыт\n" +
                                $"{externalDoc.Title}", "Информация");

                            uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                            uiapp.Application.FailuresProcessing -= OnFailureProcessing;

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
                                DisplayAllViewsInPanel(mainDocument, additionalDocument);
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при обработке документа", "Предупреждение");
                                BTN_Run.IsEnabled = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка");
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

        // Чтение видов и шаблонов
        private void DisplayAllViewsInPanel(Document mainDocument, Document additionalDocument)
        {
            SP_OtherFileSetings.Children.Clear();
         
            var viewTemplateFilter = new ElementClassFilter(typeof(View));

            mainDocumentTemplates = new SortedDictionary<string, View>();
            mainDocumentCollector = new FilteredElementCollector(mainDocument).WherePasses(viewTemplateFilter).Cast<View>();
            foreach (var view in mainDocumentCollector)
            {
                if (view.IsTemplate)
                {
                    mainDocumentTemplates[view.Name] = view;
                }          
            }

            additionalDocumentTemplates = new SortedDictionary<string, View>();
            additionalDocumentViews = new SortedDictionary<string, View>();
            additionalDocumentCollector = new FilteredElementCollector(additionalDocument).WherePasses(viewTemplateFilter).Cast<View>();
            foreach (var view in additionalDocumentCollector)
            {
                if (!view.IsTemplate)
                {
                    additionalDocumentViews[view.Name] = view;
                }
                else
                {
                    additionalDocumentTemplates[view.Name] = view;
                }
            }

            ScrollViewer scrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });

            TextBlock nameHeader = new TextBlock
            {
                Text = "Имя вида",
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(5, 3, 0, 3),
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

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(nameHeader, 0);
            Grid.SetColumn(nameHeader, 0);
            grid.Children.Add(nameHeader);
            Grid.SetRow(templateHeader, 0);
            Grid.SetColumn(templateHeader, 1);
            grid.Children.Add(templateHeader);
            int row = 1;

            foreach (var kvp in additionalDocumentViews)
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
                        isFromMain = mainDocumentTemplates.ContainsKey(initialTemplateName);
                        isFromAdditional = additionalDocumentTemplates.ContainsKey(initialTemplateName);
                    }
                }

                List<string> templateList = new List<string> { "Нет шаблона" };
                templateList.AddRange(mainDocumentTemplates.Keys);

                if (!mainDocumentTemplates.ContainsKey(initialTemplateName) &&
                    additionalDocumentTemplates.ContainsKey(initialTemplateName) &&
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

                templateCheckBox.Checked += (s, e) =>
                {
                    templateComboBox.IsEnabled = true;
                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox);
                };

                templateCheckBox.Unchecked += (s, e) =>
                {
                    templateComboBox.IsEnabled = false;

                    if (mainDocumentTemplates.ContainsKey(lockedTemplate))
                    {
                        templateComboBox.SelectedItem = lockedTemplate;
                        templateComboBox.Foreground = Brushes.Black;
                    }
                    else if (additionalDocumentTemplates.ContainsKey(lockedTemplate))
                    {
                        templateComboBox.SelectedItem = lockedTemplate;
                        templateComboBox.Foreground = Brushes.Red;
                    }
                    else
                    {
                        templateComboBox.SelectedItem = "Нет шаблона";
                        templateComboBox.Foreground = Brushes.Black;
                    }

                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox);
                };

                templateComboBox.SelectionChanged += (s, e) =>
                {
                    if (templateComboBox.SelectedItem is string selected)
                    {
                        if (selected == "Не выбрано")
                        {
                            templateComboBox.Foreground = Brushes.Black;
                        }
                        else if (mainDocumentTemplates.ContainsKey(selected))
                        {
                            templateComboBox.Foreground = Brushes.Black;
                        }
                        else
                        {
                            templateComboBox.SelectedItem = "Не выбрано";
                            templateComboBox.Foreground = Brushes.Black;
                        }
                    }

                    UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox);
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

                row++;

                UpdateTemplateChange(view.Name, templateComboBox, templateCheckBox);
            }

            scrollViewer.Content = grid;
            SP_OtherFileSetings.Children.Add(scrollViewer);
            BTN_Run.IsEnabled = true;
        }

        // Вспомогательный метод записи данных с интерфейса в словарь
        void UpdateTemplateChange(string viewName, ComboBox comboBox, CheckBox checkBox)
        {
            string selectedName = comboBox.SelectedItem as string ?? comboBox.Text;

            View selectedTemplate = null;
            mainDocumentTemplates.TryGetValue(selectedName, out selectedTemplate);

            string status;       
            if (checkBox.IsChecked == true)
            {
                if (selectedName == "Нет шаблона")
                {
                    status = "delete";
                }
                else
                {
                    status = "resave";                    
                }
            }
            else
            {
                status = "ignore";
            }

            viewTemplateChanges[viewName] = new Tuple<string, View>(status, selectedTemplate);
        }










        private void BTN_Run_Click(object sender, RoutedEventArgs e)
        {
            // ОТЛАДКА
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = System.IO.Path.Combine(desktopPath, "ViewTemplateChanges.txt");
            using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                foreach (var kvp in viewTemplateChanges)
                {
                    string key = kvp.Key;
                    string status = kvp.Value.Item1;
                    string viewName = kvp.Value.Item2 != null ? kvp.Value.Item2.Name : "Нет шаблона";
                    
                    writer.WriteLine($"{key}: {viewName} - {status}");
                }
            }
            MessageBox.Show("Файл сохранён на рабочем столе:\n" + filePath);

            // КОПИРОВАНИЕ ШАБЛОНОВ
            List<string> errorMessages = new List<string>();

            using (Transaction trans = new Transaction(additionalDocument, "KPLN. Копирование шаблонов видов"))
            {
                trans.Start();

                foreach (var kvp in viewTemplateChanges)
                {
                    string viewName = kvp.Key;
                    string status = kvp.Value.Item1;
                    View template = kvp.Value.Item2;

                    View viewToUpdate = new FilteredElementCollector(additionalDocument)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .FirstOrDefault(v => !v.IsTemplate && v.Name == viewName);

                    if (viewToUpdate == null)
                        continue;

                    if (status == "delete")
                    {
                        viewToUpdate.ViewTemplateId = ElementId.InvalidElementId;
                    }
                    else if (status == "resave" && template != null)
                    {
                        if (template.IsTemplate && viewToUpdate.ViewType == template.ViewType)
                        {
                            viewToUpdate.ViewTemplateId = template.Id;
                        }
                        else
                        {
                            errorMessages.Add($"❌ Вид \"{viewToUpdate.Name}\" не совместим с шаблоном \"{template.Name}\"");
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
