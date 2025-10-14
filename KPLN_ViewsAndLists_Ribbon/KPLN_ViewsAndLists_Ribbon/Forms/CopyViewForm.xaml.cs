using System;
using System.Linq;
using System.Windows.Media;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
        private readonly UIApplication _uiapp;
        private readonly Document _mainDocument;

        SortedDictionary<string, View> mainTemplates;
        private Dictionary<string, Tuple<string, string, string, string, View>> viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, string, View>>();

        public CopyViewForm(UIApplication uiapp)
        {
            InitializeComponent();
            _uiapp = uiapp;

            _mainDocument = uiapp.ActiveUIDocument.Document;           
            string mainDocPath = _mainDocument.PathName;
            string mainDocName = _mainDocument.Title;

            if (mainDocPath != null && mainDocName != null)
            {
                TB_ActivDocumentName.Text = $"{mainDocName}";
                TB_ActivDocumentName.ToolTip = $"{mainDocPath}";
            }

            DisplayAllTemplatesInPanel(_mainDocument);

            _copyHandler = new CopyViewExternalHandler();
            _copyEvent = ExternalEvent.Create(_copyHandler);
        }
       
        private void DisplayAllTemplatesInPanel(Document mainDocument)
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
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });

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
            TextBlock create3DHeader = new TextBlock
            {
                Text = "Создать 3D-вид",
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
            Grid.SetRow(create3DHeader, 0); Grid.SetColumn(create3DHeader, 4); grid.Children.Add(create3DHeader);

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

                CheckBox create3DCheckBox = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5),
                    ToolTip = "Создать из этого шаблона изометрический 3D-вид и применить к нему этот шаблон",
                    IsEnabled = false
                };
                bool is3DTemplate = templateView.ViewType == ViewType.ThreeD;

                selectCheckBox.Checked += (s, e) =>
                {
                    copyCheckBox.IsEnabled = true;
                    copyCheckBox.IsChecked = true;
                    replaceCheckBox.IsEnabled = true;
                    replaceCheckBox.IsChecked = true;

                    create3DCheckBox.IsEnabled = is3DTemplate; 
                    create3DCheckBox.IsChecked = false;

                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                };

                selectCheckBox.Unchecked += (s, e) =>
                {
                    replaceCheckBox.IsEnabled = false;
                    replaceCheckBox.IsChecked = false;
                    copyCheckBox.IsEnabled = false;
                    copyCheckBox.IsChecked = false;

                    create3DCheckBox.IsEnabled = false; 
                    create3DCheckBox.IsChecked = false;

                    UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                };

                replaceCheckBox.Checked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                replaceCheckBox.Unchecked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                copyCheckBox.Checked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                copyCheckBox.Unchecked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                create3DCheckBox.Checked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
                create3DCheckBox.Unchecked += (s, e) => UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);

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

                Grid.SetRow(create3DCheckBox, row);
                Grid.SetColumn(create3DCheckBox, 4);
                grid.Children.Add(create3DCheckBox);

                row++;

                UpdateTemplateBackupChoice(templateName, selectCheckBox, replaceCheckBox, copyCheckBox, create3DCheckBox, templateView);
            }

            scrollViewer.Content = grid;
            SP_OtherFileSetings.Children.Add(scrollViewer);

            BTN_RunManyFile.IsEnabled = true;
            BTN_RunManyFileRS.IsEnabled = true;
        }

        private void UpdateTemplateBackupChoice(string templateName, CheckBox selectCheckBox, CheckBox replaceCheckBox, CheckBox copyCheckBox, CheckBox create3DCheckBox, View templateView)
        {
            string statusView = selectCheckBox.IsChecked == true ? "resave" : "ignore";
            string statusResaveInV = replaceCheckBox.IsChecked == true ? "resaveIV" : "ignoreIV";
            string statusCopy = copyCheckBox.IsChecked == true ? "copyView" : "ignoreCopyView";
            string statusCreate3D = create3DCheckBox.IsChecked == true ? "create3D" : "ignoreCreate3D";

            viewOnlyTemplateChanges[templateName] =
                new Tuple<string, string, string, string, View>(statusView, statusResaveInV, statusCopy, statusCreate3D, templateView);
        }


        private void BTN_RunManyFile_Click(object sender, RoutedEventArgs e)
        {
            List<Document> openedDocs = _uiapp.Application.Documents
                .Cast<Document>()
                .Where(doc =>
                    doc.PathName != _mainDocument.PathName &&             
                    !doc.PathName.Equals(_mainDocument.PathName, StringComparison.OrdinalIgnoreCase) &&
                    !doc.IsLinked)          
                .ToList();

            if (openedDocs.Count > 0)
            {
                MessageBox.Show($"Открыто более одного документа. Для продолжения работы сохраните внесённые ранее изменения во всех документах и оставьте только документ, из которого будут копироваться шаблоны видов.", "KPLN. Информация");
                return;
            }

            if (viewOnlyTemplateChanges.Count == 0)
            {             
                MessageBox.Show($"Нет изменений для приминения", "KPLN. Информация");
                return;
            }

            var openManyDocsWindows = new ManyDocumentsSelectionWindow(_uiapp, _mainDocument, viewOnlyTemplateChanges, (bool)CHK_OpenDocuments.IsChecked, false)
            {
                Owner = this
            };
            openManyDocsWindows.ShowDialog();
        }

        private void BTN_RunManyFileRS_Click(object sender, RoutedEventArgs e)
        {
            List<Document> openedDocs = _uiapp.Application.Documents
                .Cast<Document>()
                .Where(doc =>
                    doc.PathName != _mainDocument.PathName &&
                    !doc.PathName.Equals(_mainDocument.PathName, StringComparison.OrdinalIgnoreCase) &&
                    !doc.IsLinked)
                .ToList();

            if (openedDocs.Count > 0)
            {
                MessageBox.Show($"Открыто более одного документа. Для продолжения работы сохраните внесённые ранее изменения во всех документах и оставьте только документ, из которого будут копироваться шаблоны видов.", "KPLN. Информация");
                return;
            }

            if (viewOnlyTemplateChanges.Count == 0)
            {
                MessageBox.Show($"Нет изменений для приминения", "KPLN. Информация");
                return;
            }


            var openManyDocsWindows = new ManyDocumentsSelectionWindow(_uiapp, _mainDocument, viewOnlyTemplateChanges, (bool)CHK_OpenDocuments.IsChecked, true)
            {
                Owner = this
            };
            openManyDocsWindows.ShowDialog();
        }
    }
}