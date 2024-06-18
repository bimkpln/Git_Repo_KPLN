using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
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

        // Обработчик ошибок и предупреждений
        public class IgnoreFailuresPreprocessor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
                foreach (FailureMessageAccessor failure in failures)
                {
                    failuresAccessor.DeleteWarning(failure);
                }
                return FailureProcessingResult.Continue;
            }
        }

        // Создаём словарь: key - string (level name); value - List<ElementId>> 
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

        // XAML: обработчик события LevelExport
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

        // XAML: обработчик события LevelImport
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

        // XAML: обновление содержимого списков LevelExport и LevelImport
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

        // Формула сдвига элемента на уровне
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

        // Создать и открыть новый вид
        private View3D CreateAndOpenNewView(string exportLevelName, string importLevelName)
        {
            ViewFamilyType viewFamilyType = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);          

            if (viewFamilyType != null)
            {
                View3D newView3D = null;

                using (Transaction transaction = new Transaction(_doc, "KPLN: Создание нового 3D вида"))
                {
                    transaction.Start();
                   
                    // Используем стандартные 3D настройки (изометрию)
                    XYZ eyePosition = new XYZ(1, 1, 1);
                    XYZ upDirection = new XYZ(0, 0, 1);
                    XYZ forwardDirection = new XYZ(0, 1, 0);

                    ViewOrientation3D viewOrientation = new ViewOrientation3D(eyePosition, upDirection, forwardDirection);

                    newView3D = View3D.CreateIsometric(_doc, viewFamilyType.Id);

                    if (newView3D != null)
                    {
                        List<string> existingViewNames = new List<string>();

                        FilteredElementCollector collector = new FilteredElementCollector(_doc);
                        ICollection<Element> views = collector.OfClass(typeof(View)).ToElements();

                        foreach (Element view in views)
                        {
                            View viewElement = view as View;

                            if (viewElement != null && !string.IsNullOrEmpty(viewElement.Name))
                            {
                                existingViewNames.Add(viewElement.Name);
                            }
                        }

                        string baseName = $"ПроверочныйВид__{exportLevelName}--{importLevelName}";
                        string newViewName = baseName;
                        int numSuffixView3D = 2;

                        while (existingViewNames.Contains(newViewName))
                        {
                            newViewName = $"{baseName} ({numSuffixView3D++})";
                        }

                        newView3D.Name = newViewName;

                        newView3D.SetOrientation(viewOrientation);
                    }

                    transaction.Commit();
                }

                if (newView3D != null)
                {
                    UIDocument uiDoc = new UIDocument(_doc);
                    uiDoc.ActiveView = newView3D;
                    return newView3D;
                }
            }

            return null;
        }

        // Условия для окон с предупреждениями
        private bool WarningDialogWindow(string exportLevelName, string importLevelName)
        {
            if (exportLevelName == importLevelName)
            {
                System.Windows.Forms.MessageBox.Show("Выбраны одинаковые уровни", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            if (LevelExportElementList.Text == "")
            {
                System.Windows.Forms.MessageBox.Show("На уровне-экспортере нет элементов", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // Перемещение элементов на новый уровень
        private void Button_ClickTransferringElements(object sender, RoutedEventArgs e)
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;

            bool shouldContinue = WarningDialogWindow(exportLevelName, importLevelName);

            if (!shouldContinue)
            {
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

                    // Разгруппировка эллементов
                    Group groupOld = element as Group;
                    if (groupOld != null)
                    {
                        groupOld.UngroupMembers();
                    }

                    //
                    var levelNameParameter = element.get_Parameter(levelExportParametrs[0]);
                    var levelOffsetParameter = element.get_Parameter(levelExportParametrs[1]);

                    if (levelNameParameter?.HasValue == true && levelNameParameter.IsReadOnly == false)
                    {
                        FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();
                        failureHandlingOptions.SetFailuresPreprocessor(new IgnoreFailuresPreprocessor());
                        transaction.SetFailureHandlingOptions(failureHandlingOptions);

                        levelOffsetParameter.Set(CalculatedElementOffset(element, newLevelTransfer));
                        levelNameParameter.Set(newLevel.Id);                                               
                    }
                }

                //Группировка эллементов
                Group groupNew = _doc.Create.NewGroup(ElementsByLevel[exportLevelName]);

                transaction.Commit();
            }

            CreateAndOpenNewView(exportLevelName, importLevelName);

            ElementsByLevel = GetElementsByLevel(_doc);
            ElementLevelListName();               
        }

        // XAML: Переместить на уровень отсоеденив зависимость сверху
        private void Button_ClickTransferringElementsConditions(object sender, RoutedEventArgs e)
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;

            bool shouldContinue = WarningDialogWindow(exportLevelName, importLevelName);

            if (!shouldContinue)
            {
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

                    // Разгруппировка эллементов
                    Group groupOld = element as Group;
                    if (groupOld != null)
                    {
                        groupOld.UngroupMembers();
                    }

                    //
                    var levelNameParameter = element.get_Parameter(levelExportParametrs[0]);
                    var levelOffsetParameter = element.get_Parameter(levelExportParametrs[1]);

                    if (levelNameParameter?.HasValue == true && levelNameParameter.IsReadOnly == false)
                    {
                        FailureHandlingOptions failureHandlingOptions = transaction.GetFailureHandlingOptions();
                        failureHandlingOptions.SetFailuresPreprocessor(new IgnoreFailuresPreprocessor());
                        transaction.SetFailureHandlingOptions(failureHandlingOptions);

                        levelOffsetParameter.Set(CalculatedElementOffset(element, newLevelTransfer));
                        levelNameParameter.Set(newLevel.Id);
                    }
                }

                //Группировка эллементов
                Group groupNew = _doc.Create.NewGroup(ElementsByLevel[exportLevelName]);

                transaction.Commit();
            }

            CreateAndOpenNewView(exportLevelName, importLevelName);

            ElementsByLevel = GetElementsByLevel(_doc);
            ElementLevelListName();
        }

        // XAML: закрыть окно
        private void Button_CloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
