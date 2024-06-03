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
        public Dictionary<string, List<ElementId>> ElementsByLevel { get; set; }

        public FormChageLevel(Document document)
        {
            _doc = document;
            ElementsByLevel = GetElementsByLevel(_doc);

            InitializeComponent();
            this.DataContext = this;
        }

        public Dictionary<string, List<ElementId>> GetElementsByLevel(Document doc)
        {
            Dictionary<string, List<ElementId>> elementsByLevel = new Dictionary<string, List<ElementId>>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);

            List<Level> levels = collector.OfClass(typeof(Level)).Cast<Level>().ToList();

            foreach (Level level in levels)
            {
                List<ElementId> elementInLevel = new List<ElementId>();

                ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);

                FilteredElementCollector elementsList = new FilteredElementCollector(doc);

                List<Element> elements = elementsList.WherePasses(levelFilter).ToList();

                foreach (Element element in elements)
                {
                    elementInLevel.Add(element.Id);
                }

                elementsByLevel.Add(level.Name, elementInLevel);
            }

            return elementsByLevel;
        }

        public double GetOffset(ElementId elementID, Level level)
        {
            double offset = 0;

            Element element = _doc.GetElement(elementID);  
            BuiltInParameter[] levelExportParametrs = CommandChangeLevel.GetParametersForMovingItems(element);

            double baseOffset = element.get_Parameter(levelExportParametrs[1]).AsDouble();

            Level baseLevel = _doc.GetElement(element.LevelId) as Level;
            double baseElevation = baseLevel.Elevation;

            double newElevation = level.Elevation;

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
                List<ElementId> elements = ElementsByLevel[selectedKey];

                LevelExportElementList.Text = "";

                foreach (ElementId elementID in elements)
                {
                    Element element = _doc.GetElement(elementID);
                    LevelExportElementList.Text += $"ID: {elementID}; ИМЯ: {element.Name}\n";
                }
            }
        }

        private void LevelImport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedKey = levelImport.SelectedItem as string;

            if (selectedKey != null && ElementsByLevel.ContainsKey(selectedKey))
            {
                List<ElementId> elements = ElementsByLevel[selectedKey];

                LevelImportElementList.Text = "";

                foreach (ElementId elementID in elements)
                {
                    Element element = _doc.GetElement(elementID);
                    LevelImportElementList.Text += $"ID: {elementID}; ИМЯ: {element.Name}\n";
                }
            }
        }

        private void ElementLevelListName()
        {
            LevelExportElementList.Text = "";
            LevelImportElementList.Text = "";

            foreach (ElementId elementID in ElementsByLevel[LevelExport.SelectedItem as string])
            {
                Element element = _doc.GetElement(elementID);
                LevelExportElementList.Text += $"ID: {elementID}; ИМЯ: {element.Name}\n";
            }

            foreach (ElementId elementID in ElementsByLevel[levelImport.SelectedItem as string])
            {
                Element element = _doc.GetElement(elementID);
                LevelImportElementList.Text += $"ID: {elementID}; ИМЯ: {element.Name}\n";
            }
        }

        private double CalculatedElementOffset(Element element, Level level)
        {
            double offset = 0;

            BuiltInParameter[] levelExportParametrs = CommandChangeLevel.GetParametersForMovingItems(element);
            double baseOffset = element.get_Parameter(levelExportParametrs[1]).AsDouble();

            Level baseLevel = _doc.GetElement(element.LevelId) as Level;
            double baseElevation = baseLevel.Elevation;
            double newElevation = level.Elevation;

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

        private void Button_CloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_ClickTransferringElements(object sender, RoutedEventArgs e)
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;

            if (exportLevelName == importLevelName)
            {
                System.Windows.Forms.MessageBox.Show("Выбраны одинаковые уровни", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            if (LevelExportElementList.Text == "")
            {
                System.Windows.Forms.MessageBox.Show("На уровне-экспортере нет элементов", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            Level newLevel = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level)).FirstOrDefault(x => x.Name == importLevelName) as Level;
            Level newLevelTransfer = newLevel;

            using (Transaction transaction = new Transaction(_doc, "KPLN: Перенос на новый уровень"))
            {
                transaction.Start();
             
                foreach (ElementId elementID in ElementsByLevel[exportLevelName])
                {
                    Element element = _doc.GetElement(elementID);
                    BuiltInParameter[] levelExportParametrs = CommandChangeLevel.GetParametersForMovingItems(element);

                    var levelNameParameter = element.get_Parameter(levelExportParametrs[0]);
                    var levelOffsetParameter = element.get_Parameter(levelExportParametrs[1]);

                    if (levelNameParameter?.HasValue == true && levelNameParameter.IsReadOnly == false && levelOffsetParameter.IsReadOnly == false)
                    {                       
                        levelOffsetParameter.Set(CalculatedElementOffset(element, newLevelTransfer));
                        levelNameParameter.Set(newLevel.Id);
                    }
                }

                transaction.Commit();
            }

            ElementsByLevel = GetElementsByLevel(_doc);
            ElementLevelListName();
        }
    }
}
