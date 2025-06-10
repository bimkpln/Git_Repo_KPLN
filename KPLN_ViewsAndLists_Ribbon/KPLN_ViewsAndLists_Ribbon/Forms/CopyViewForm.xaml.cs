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

        UIApplication _uiapp;

        Document mainDocument = null;

        SortedDictionary<string, View> mainTemplates;
        private Dictionary<string, Tuple<string, string, string, View>> viewOnlyTemplateChanges = new Dictionary<string, Tuple<string, string, string, View>>();

        public CopyViewForm(UIApplication uiapp, Document heritageMainDocument, Document heritageAdditionalDocument)
        {
            InitializeComponent();
            _uiapp = uiapp;

            mainDocument = uiapp.ActiveUIDocument.Document;           
            string mainDocPath = mainDocument.PathName;
            string mainDocName = mainDocument.Title;

            if (mainDocPath != null && mainDocName != null)
            {
                TB_ActivDocumentName.Text = $"{mainDocName}";
                TB_ActivDocumentName.ToolTip = $"{mainDocPath}";
            }

            DisplayAllTemplatesInPanel(mainDocument);

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

            BTN_RunManyFile.IsEnabled = true;
            BTN_RunManyFileRS.IsEnabled = true;
        }

        private void UpdateTemplateBackupChoice(string templateName, CheckBox selectCheckBox, CheckBox replaceCheckBox, CheckBox copyCheckBox, View templateView)
        {
            string statusView = selectCheckBox.IsChecked == true ? "resave" : "ignore";
            string statusResaveInView = replaceCheckBox.IsChecked == true ? "resaveIV" : "ignoreIV";
            string statusCopy = copyCheckBox.IsChecked == true ? "copyView" : "ignoreCopyView";

            viewOnlyTemplateChanges[templateName] = new Tuple<string, string, string, View>(statusView, statusResaveInView, statusCopy, templateView);
        }
   
        private void BTN_RunManyFile_Click(object sender, RoutedEventArgs e)
        {
            List<Document> openedDocs = _uiapp.Application.Documents
                .Cast<Document>()
                .Where(doc =>
                    doc.PathName != mainDocument.PathName &&             
                    !doc.PathName.Equals(mainDocument.PathName, StringComparison.OrdinalIgnoreCase) &&
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
        
            var openManeDocsWindows = new ManyDocumentsSelectionWindow(_uiapp, mainDocument, viewOnlyTemplateChanges, false);
            openManeDocsWindows.Owner = this;       
            openManeDocsWindows.ShowDialog();
        }

        private void BTN_RunManyFileRS_Click(object sender, RoutedEventArgs e)
        {
            List<Document> openedDocs = _uiapp.Application.Documents
                .Cast<Document>()
                .Where(doc =>
                    doc.PathName != mainDocument.PathName &&
                    !doc.PathName.Equals(mainDocument.PathName, StringComparison.OrdinalIgnoreCase) &&
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

            var openManeDocsWindows = new ManyDocumentsSelectionWindow(_uiapp, mainDocument, viewOnlyTemplateChanges, true);
            openManeDocsWindows.Owner = this;
            openManeDocsWindows.ShowDialog();
        }
    }
}