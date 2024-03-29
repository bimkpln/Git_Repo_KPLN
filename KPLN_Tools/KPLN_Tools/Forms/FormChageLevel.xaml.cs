using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using KPLN_Tools.ExternalCommands;


namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormChageLevel.xaml
    /// </summary>
    public partial class FormChageLevel : Window
    {
        private Document _doc;
        public Dictionary<string, List<Element>> ElementsByLevel { get; set; }
        public string[] LevelNames { get; set; }

        public FormChageLevel(Document document)
        {
            _doc = document;

            InitializeComponent();

            ElementsByLevel = GetElementsByLevel(_doc);

            this.DataContext = this;
        }

        public Dictionary<string, List<Element>> GetElementsByLevel(Document doc)
        {
            Dictionary<string, List<Element>> elementsByLevel = new Dictionary<string, List<Element>>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Level> levels = collector.OfClass(typeof(Level)).Cast<Level>().ToList();

            foreach (Level level in levels)
            {
                List<Element> elementNames = new List<Element>();

                ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);

                FilteredElementCollector elementsList = new FilteredElementCollector(doc);
                List<Element> elements = elementsList.WherePasses(levelFilter).ToList();

                foreach (Element element in elements)
                {
                    elementNames.Add(element);
                }

                elementsByLevel.Add(level.Name, elementNames);
            }

            return elementsByLevel;
        }

        public void TransferringElementsToAnotherLevel()
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;

            foreach (Element element in ElementsByLevel[exportLevelName])
            {
                
            }
        }

        public double GetOffset(Element element)
        {
            double offset = 0;

            BuiltInParameter[] levelExportParametrs = CommandChangeLevel.GetParametersForMovingItems(element);

            double baseOffset = element.get_Parameter(levelExportParametrs[1]).AsDouble();

            Level baseLevel = _doc.GetElement(element.LevelId) as Level;

            double baseElevation = baseLevel.Elevation;
            double newElevation = baseLevel.Elevation;
            double elevationDifference = newElevation - baseElevation;

            if (elevationDifference > 0)
            {
                offset = baseOffset - Math.Abs(elevationDifference);
            }
            if (elevationDifference < 0)
            {
                offset = baseOffset + Math.Abs(elevationDifference);
            }

            return offset;
        }

        private void LevelExport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedKey = LevelExport.SelectedItem as string;

            if (selectedKey != null && ElementsByLevel.ContainsKey(selectedKey))
            {
                List<Element> elements = ElementsByLevel[selectedKey];

                LevelExportElementList.Text = "";

                foreach (Element element in elements)
                {
                    LevelExportElementList.Text += $"ID: {element.Id}; ИМЯ: {element.Name}\n";
                }
            }
        }

        private void LevelImport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedKey = levelImport.SelectedItem as string;

            if (selectedKey != null && ElementsByLevel.ContainsKey(selectedKey))
            {
                List<Element> elements = ElementsByLevel[selectedKey];

                LevelImportElementList.Text = "";

                foreach (Element element in elements)
                {
                    LevelImportElementList.Text += $"ID: {element.Id}; ИМЯ: {element.Name}\n";
                }
            }
        }

        private void Button_CloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_ClickTransferringElements(object sender, RoutedEventArgs e)
        {
            TransferringElementsToAnotherLevel();
        }
    }
}
