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
using System.Xml.Linq;
using System.Windows.Documents;
using static System.Net.Mime.MediaTypeNames;
using System.Windows.Forms;


namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormChageLevel.xaml
    /// </summary>
    public partial class FormChageLevel : Window
    {
        private Document _doc;
        public Dictionary<string, List<ElementId>> ElementsByLevel { get; set; }

        public static string listLevelExportElement;
        public static string listLevelImportElement;

        bool conditionsTopLevel;

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

            listLevelExportElement = "";

            if (selectedKey != null && ElementsByLevel.ContainsKey(selectedKey))
            {
                List<ElementId> elements = ElementsByLevel[selectedKey];
           
                foreach (ElementId elementID in elements)
                {
                    Element element = _doc.GetElement(elementID);
                    listLevelExportElement += $"ID: {elementID}; ИМЯ: {element.Name}\n";
                }

                System.Windows.Controls.RichTextBox levelExportElementListRTB = LevelExportElementList;
                levelExportElementListRTB.Document.Blocks.Clear();

                Paragraph paragraphLEChanges = new Paragraph();
                paragraphLEChanges.Inlines.Add(new Run(listLevelExportElement));
                LevelExportElementList.Document.Blocks.Add(paragraphLEChanges);
            }
        }

        // XAML: обработчик события LevelImport
        private void LevelImport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedKey = levelImport.SelectedItem as string;

            listLevelImportElement = "";

            if (selectedKey != null && ElementsByLevel.ContainsKey(selectedKey))
            {
                List<ElementId> elements = ElementsByLevel[selectedKey];              

                foreach (ElementId elementID in elements)
                {
                    Element element = _doc.GetElement(elementID);
                    listLevelImportElement += $"ID: {elementID}; ИМЯ: {element.Name}\n";
                }

                LevelImportElementList.Document.Blocks.Clear();
                Paragraph paragraphLIChanges = new Paragraph();
                paragraphLIChanges.Inlines.Add(new Run(listLevelImportElement));
                LevelImportElementList.Document.Blocks.Add(paragraphLIChanges);
            }
        }

        // XAML: обновление содержимого списков LevelExport и LevelImport
        private void ElementLevelListName()
        {
            listLevelExportElement = "";
            listLevelImportElement = "";

            foreach (ElementId elementID in ElementsByLevel[LevelExport.SelectedItem as string])
            {
                Element element = _doc.GetElement(elementID);
                listLevelExportElement += $"ID: {elementID}; ИМЯ: {element.Name}\n";
            }

            foreach (ElementId elementID in ElementsByLevel[levelImport.SelectedItem as string])
            {
                Element element = _doc.GetElement(elementID);
                listLevelImportElement += $"ID: {elementID}; ИМЯ: {element.Name}\n";
            }

            LevelExportElementList.Document.Blocks.Clear();
            LevelImportElementList.Document.Blocks.Clear();

            Paragraph paragraphExport = new Paragraph();
            paragraphExport.Inlines.Add(new Run(listLevelExportElement));
            LevelExportElementList.Document.Blocks.Add(paragraphExport);

            Paragraph paragraphImport = new Paragraph();
            paragraphImport.Inlines.Add(new Run(listLevelImportElement));
            LevelImportElementList.Document.Blocks.Add(paragraphImport);
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
                .OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

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
                        ICollection<Element> views = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

                        foreach (Element view in views)
                        {
                            Autodesk.Revit.DB.View viewElement = view as Autodesk.Revit.DB.View;

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

                    ApplyTemporaryIsolation(newView3D, importLevelName);

                    return newView3D;
                }
            }

            return null;
        }

        // Временаая изоляция элементов
        private void ApplyTemporaryIsolation(View3D view3D, string levelName)
        {
            using (Transaction transaction = new Transaction(_doc, "KPLN: Временная изоляция элементов"))
            {
                transaction.Start();

                Level level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .FirstOrDefault(e => e.Name.Equals(levelName)) as Level;

                if (level != null)
                {
                    ElementLevelFilter levelFilter = new ElementLevelFilter(level.Id);
                    ICollection<ElementId> elementIds = new FilteredElementCollector(_doc)
                        .WherePasses(levelFilter)
                        .WhereElementIsNotElementType()
                        .ToElementIds();

                    if (elementIds.Count > 0)
                    {
                        view3D.IsolateElementsTemporary(elementIds);
                    }
                }

                transaction.Commit();
            }
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

            if (string.IsNullOrEmpty(new System.Windows.Documents.TextRange(LevelExportElementList.Document.ContentStart, LevelExportElementList.Document.ContentEnd).Text))
            {
                System.Windows.Forms.MessageBox.Show("На уровне-экспортере нет элементов", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            if (NewTopLevel.Text == "" && conditionsTopLevel == true)
            {
                System.Windows.Forms.MessageBox.Show("Не выбран верхний уровень", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            if (NewTopLevel.Text == exportLevelName && conditionsTopLevel == true)
            {
                System.Windows.Forms.MessageBox.Show("Выбранный верхний уровень совпадает с уровнем-экспортером", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            if (NewTopLevel.Text == importLevelName && conditionsTopLevel == true)
            {
                System.Windows.Forms.MessageBox.Show("Выбранный верхний уровень совпадает с уровнем-импортером", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // Окно о выполнении работы
        private void FinishTransferMessage()
        {
            if (!string.IsNullOrEmpty(new System.Windows.Documents.TextRange(LevelImportElementList.Document.ContentStart, LevelImportElementList.Document.ContentEnd).Text)) 
            {
                System.Windows.Forms.MessageBox.Show("Не все элементы были перенесены на новый уровень", "Предупреждение",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Все элементы были перенесены на новый уровень", "Уведомление",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }
        }    

        // Нахождение высоты черезз геометрию (стены, колоны)
        private double GetElementHeight(Element element)
        {
            Options options = new Options();
            GeometryElement geomElement = element.get_Geometry(options);

            double minZ = double.MaxValue;
            double maxZ = double.MinValue;

            foreach (GeometryObject geomObj in geomElement)
            {
                Solid solid = geomObj as Solid;
                if (solid != null)
                {
                    foreach (Face face in solid.Faces)
                    {
                        Mesh mesh = face.Triangulate();
                        foreach (XYZ vertex in mesh.Vertices)
                        {
                            if (vertex.Z < minZ) minZ = vertex.Z;
                            if (vertex.Z > maxZ) maxZ = vertex.Z;
                        }
                    }
                }
            }

            double height = maxZ - minZ;
            return height;
            }
     
        // Перемещение элементов на новый уровень
        private void Button_ClickTransferringElements(object sender, RoutedEventArgs e)
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;

            conditionsTopLevel = false;

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
                try
                {
                    Group groupNew = _doc.Create.NewGroup(ElementsByLevel[exportLevelName]);
                } catch (Autodesk.Revit.Exceptions.ArgumentException) { };

                transaction.Commit();
            }

            CreateAndOpenNewView(exportLevelName, importLevelName);

            ElementsByLevel = GetElementsByLevel(_doc);
            ElementLevelListName();

            FinishTransferMessage();
        }

        // Перемещение элементов на новый уровень, отсоеденив зависимость сверху
        private void Button_ClickTransferringElementsConditions(object sender, RoutedEventArgs e)
        {
            string exportLevelName = LevelExport.SelectedItem as string;
            string importLevelName = levelImport.SelectedItem as string;
            string newTopLevellName = NewTopLevel.SelectedItem as string;

            conditionsTopLevel = true;

            bool shouldContinue = WarningDialogWindow(exportLevelName, importLevelName);

            if (!shouldContinue)
            {
                return;
            }

            Level newLevel = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level)).FirstOrDefault(x => x.Name == importLevelName) as Level;
            Level newLevelTransfer = newLevel;

            Level newTopLevel = new FilteredElementCollector(_doc)
               .OfClass(typeof(Level)).FirstOrDefault(x => x.Name == newTopLevellName) as Level;

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

                    if (levelExportParametrs.Length > 2 && (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls || 
                        element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns))
                    {
                        var topLevelNameParameter = element.get_Parameter(levelExportParametrs[2]);
                        var topLevelOffsetParameter = element.get_Parameter(levelExportParametrs[3]);
                        var elementHeight = element.get_Parameter(levelExportParametrs[4]);

                        double heightValue = GetElementHeight(element);

                        topLevelNameParameter.Set(ElementId.InvalidElementId);
                        elementHeight.Set(heightValue);                       
                    }

                    if (levelExportParametrs.Length > 2 && (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Stairs 
                        || element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Ramps))
                    {
                        var topLevelNameParameter = element.get_Parameter(levelExportParametrs[2]);
                        var topLevelOffsetParameter = element.get_Parameter(levelExportParametrs[3]);

                        levelNameParameter.Set(newLevel.Id);
                        topLevelNameParameter.Set(newTopLevel.Id) ;
                    }

                }

                //Группировка эллементов
                try
                {
                    Group groupNew = _doc.Create.NewGroup(ElementsByLevel[exportLevelName]);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException) { };

                transaction.Commit();
            }

            CreateAndOpenNewView(exportLevelName, importLevelName);

            ElementsByLevel = GetElementsByLevel(_doc);
            ElementLevelListName();

            FinishTransferMessage();
        }

        // XAML: закрыть окно
        private void Button_CloseClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
