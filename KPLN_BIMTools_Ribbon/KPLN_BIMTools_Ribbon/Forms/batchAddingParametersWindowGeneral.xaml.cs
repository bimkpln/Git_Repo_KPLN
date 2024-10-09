using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Input;
using static System.Net.Mime.MediaTypeNames;
using System.Diagnostics;
using System.Windows.Data;


namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowGeneral : Window
    {
        public UIApplication uiapp;
        public Autodesk.Revit.ApplicationServices.Application revitApp;
        public Document _doc;

        public string activeFamilyName;
        public string jsonFileSettingPath;

        public string SPFPath;
        public Dictionary<String, List<ExternalDefinition>> groupAndParametersFromSPFDict = new Dictionary<String, List<ExternalDefinition>>();
        List<string> allParamNameList = new List<string>();

        public batchAddingParametersWindowGeneral(UIApplication uiapp, string activeFamilyName, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            revitApp = uiapp.Application;
            _doc = uiapp.ActiveUIDocument?.Document;

            this.activeFamilyName = activeFamilyName;
            this.jsonFileSettingPath = jsonFileSettingPath;

            TB_familyName.Text = activeFamilyName;

            FillingComboBoxTypeInstance();
            FillingComboBoxGroupingName();

            if (!string.IsNullOrEmpty(jsonFileSettingPath))
            {
                StackPanel parentContainer = (StackPanel)SP_panelParamFields.Parent;
                parentContainer.Children.Remove(SP_panelParamFields);

                Dictionary<string, List<string>> allParamInInterfaceFromJsonDict = ParsingDataFromJsonToSPF(jsonFileSettingPath);

                AddPanelParamFieldsJson(allParamInInterfaceFromJsonDict);
            }
            else
            {
                SPFPath = revitApp.SharedParametersFilename;

                HandlerGeneralParametersFile(SPFPath);

                TB_filePath.Text = SPFPath;
                SP_panelParamFields.ToolTip = $"ФОП: {SPFPath}";        
                CB_paramsGroup.Tag = SPFPath;               
                CB_paramsGroup.ToolTip = $"ФОП: {SPFPath}";
            }
        }

        // Создание List всех параметров из ФОП
        static public List<string> CreateallParamNameList(Dictionary<String, List<ExternalDefinition>> groupAndParametersFromSPFDict)
        {
            List<string> allParamNameList = new List<string>();

            foreach (var kvp in groupAndParametersFromSPFDict)
            {
                foreach (var param in kvp.Value)
                {
                    allParamNameList.Add(param.Name);
                }
            }

            allParamNameList.Sort();

            return allParamNameList;
        }

        // Создание List для "Тип/Экземпляр"
        static public List<string> CreateTypeInstanceList()
        {
            List<string> typeInstanceList = new List<string>();
            typeInstanceList.Add("Тип");
            typeInstanceList.Add("Экземпляр");

            return typeInstanceList;
        }

        // Заполнение оригинального ComboBox "Тип/Экземпляр"
        public void FillingComboBoxTypeInstance()
        {
            foreach (string key in CreateTypeInstanceList())
            {
                CB_typeInstance.Items.Add(key);
            }
        }

        // Создание Dictionary с параметрами группирования для "Параметры группирования"
        static public Dictionary<string, BuiltInParameterGroup> CreateGroupingDictionary()
        {
            Dictionary<string, BuiltInParameterGroup> groupingDict = new Dictionary<string, BuiltInParameterGroup>();

            groupingDict.Add("Аналитическая модель", BuiltInParameterGroup.PG_ANALYTICAL_MODEL);
            groupingDict.Add("Видимость", BuiltInParameterGroup.PG_VISIBILITY);
            groupingDict.Add("Второстепенный конец", BuiltInParameterGroup.PG_SECONDARY_END);
            groupingDict.Add("Выравнивание аналитической модели", BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT);
            groupingDict.Add("Геометрия разделения", BuiltInParameterGroup.PG_DIVISION_GEOMETRY);
            groupingDict.Add("Графика", BuiltInParameterGroup.PG_GRAPHICS);
            groupingDict.Add("Данные", BuiltInParameterGroup.PG_DATA);
            groupingDict.Add("Зависимости", BuiltInParameterGroup.PG_CONSTRAINTS);
            groupingDict.Add("Идентификация", BuiltInParameterGroup.PG_IDENTITY_DATA);
            groupingDict.Add("Материалы и отделка", BuiltInParameterGroup.PG_MATERIALS);
            groupingDict.Add("Механизмы", BuiltInParameterGroup.PG_MECHANICAL);
            groupingDict.Add("Механизмы - Нагрузки", BuiltInParameterGroup.PG_MECHANICAL_LOADS);
            groupingDict.Add("Механизмы - Расход", BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW);
            groupingDict.Add("Моменты", BuiltInParameterGroup.PG_MOMENTS);
            groupingDict.Add("Набор", BuiltInParameterGroup.PG_COUPLER_ARRAY);
            groupingDict.Add("Набор арматурных стержней", BuiltInParameterGroup.PG_REBAR_ARRAY);
            groupingDict.Add("Несущие конструкции", BuiltInParameterGroup.PG_STRUCTURAL);
            groupingDict.Add("Общая легенда", BuiltInParameterGroup.PG_OVERALL_LEGEND);
            groupingDict.Add("Общие", BuiltInParameterGroup.PG_GENERAL);
            groupingDict.Add("Основной конец", BuiltInParameterGroup.PG_PRIMARY_END);
            groupingDict.Add("Параметры IFC", BuiltInParameterGroup.PG_IFC);
            groupingDict.Add("Прочее", BuiltInParameterGroup.INVALID);
            groupingDict.Add("Размеры", BuiltInParameterGroup.PG_GEOMETRY);
            groupingDict.Add("Расчет несущих конструкций", BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS);
            groupingDict.Add("Расчет энергопотребления", BuiltInParameterGroup.PG_ENERGY_ANALYSIS);
            groupingDict.Add("Редактирование формы перекрытия", BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT);
            groupingDict.Add("Результат анализа", BuiltInParameterGroup.PG_ANALYSIS_RESULTS);
            groupingDict.Add("Сантехника", BuiltInParameterGroup.PG_PLUMBING);
            groupingDict.Add("Свойства модели", BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES);
            groupingDict.Add("Свойства экологически чистого здания", BuiltInParameterGroup.PG_GREEN_BUILDING);
            groupingDict.Add("Сегменты и соединительные детали", BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
            groupingDict.Add("Силы", BuiltInParameterGroup.PG_FORCES);
            groupingDict.Add("Система пожаротушения", BuiltInParameterGroup.PG_FIRE_PROTECTION);
            groupingDict.Add("Слои", BuiltInParameterGroup.PG_REBAR_SYSTEM_LAYERS);
            groupingDict.Add("Снятие связей/усилия для элемента", BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES);
            groupingDict.Add("Стадии", BuiltInParameterGroup.PG_PHASING);
            groupingDict.Add("Строительство", BuiltInParameterGroup.PG_CONSTRUCTION);
            groupingDict.Add("Текст", BuiltInParameterGroup.PG_TEXT);
            groupingDict.Add("Фотометрические", BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS);
            groupingDict.Add("Шрифт заголовков", BuiltInParameterGroup.PG_TITLE);
            groupingDict.Add("Электросети (PG_ELECTRICAL_ENGINEERING)", BuiltInParameterGroup.PG_ELECTRICAL_ENGINEERING);
            groupingDict.Add("Электросети (PG_ELECTRICAL)", BuiltInParameterGroup.PG_ELECTRICAL);
            groupingDict.Add("Электросети - Нагрузки", BuiltInParameterGroup.PG_ELECTRICAL_LOADS);
            groupingDict.Add("Электросети - Освещение", BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING);
            groupingDict.Add("Электросети - Создание цепей", BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING);

            return groupingDict;
        }

        // Заполнение оригинального ComboBox "Группирование"
        public void FillingComboBoxGroupingName()
        {
            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                CB_grouping.Items.Add(groupName);
            }
        }

        // Формирование словаря с параметрами из JSON-файла преднастроек
        public Dictionary<string, List<string>> ParsingDataFromJsonToSPF(string jsonFileSettingPath)
        {
            string jsonData;

            using (StreamReader reader = new StreamReader(jsonFileSettingPath))
            {
                jsonData = reader.ReadToEnd();
            }

            var paramListFromJson = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonData);
            var newParamList = new Dictionary<string, List<string>>();

            foreach (var entry in paramListFromJson)
            {
                string nameParameter = entry["NE"];

                if (!newParamList.ContainsKey(nameParameter))
                {
                    newParamList[nameParameter] = new List<string>();
                }

                newParamList[nameParameter].Add(entry["pathFile"]);
                newParamList[nameParameter].Add(entry["groupParameter"]);
                newParamList[nameParameter].Add(entry["nameParameter"]);
                newParamList[nameParameter].Add(entry["instance"]);
                newParamList[nameParameter].Add(entry["grouping"]);
            }

            return newParamList;
        }
    
        // Обработчик ФОПа; 
        public void HandlerGeneralParametersFile(string filePath)
        {
            groupAndParametersFromSPFDict.Clear();
            revitApp.SharedParametersFilename = SPFPath;

            try
            {
                DefinitionFile defFile = revitApp.OpenSharedParameterFile();

                if (defFile == null)
                {
                    System.Windows.Forms.MessageBox.Show($"{SPFPath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    SPFPath = "";
                    TB_filePath.Text = SPFPath;

                    return;
                }

                foreach (DefinitionGroup group in defFile.Groups)
                {
                    List<ExternalDefinition> parametersList = new List<ExternalDefinition>();

                    foreach (ExternalDefinition definition in group.Definitions)
                    {
                        parametersList.Add(definition);
                    }

                    groupAndParametersFromSPFDict[group.Name] = parametersList;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"{SPFPath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                SPFPath = "";
                TB_filePath.Text = SPFPath;

                return;
            }

            allParamNameList = CreateallParamNameList(groupAndParametersFromSPFDict);

            foreach (var key in groupAndParametersFromSPFDict.Keys)
            {
                CB_paramsGroup.Items.Add(key);
            }
          
            foreach (var param in allParamNameList)
            {
                CB_paramsName.Items.Add(param);
            }
        }

        // Создание словаря со всеми параметрами указанными в интерфейсе
        public Dictionary<string, List<string>> CreateInterfaceParamDict()
        {

            Dictionary<string, List<string>> interfaceParamDict = new Dictionary<string, List<string>>();

            foreach (var stackPanel in FindVisualChildren<StackPanel>(this))
            {
                if (stackPanel.Tag != null && stackPanel.Tag.ToString() == "uniqueParameterField")
                {
                    int stackPanelIndex = interfaceParamDict.Count + 1;
                    string stackPanelKey = $"ParamFromSPFAdd-{stackPanelIndex}";
                 
                    var comboBoxes = stackPanel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();
                    List<string> comboBoxValues = new List<string>();

                    if (comboBoxes.Count >= 4)
                    {
                        string firstComboBoxTag = comboBoxes[0].Tag?.ToString() ?? "";
                        comboBoxValues.Add(firstComboBoxTag);

                        foreach (var comboBox in comboBoxes.Take(4))
                        {
                            comboBoxValues.Add(comboBox.SelectedItem?.ToString() ?? "");
                        }

                        interfaceParamDict.Add(stackPanelKey, comboBoxValues);
                    }
                }
            }

            // Проверка на дубликаты параметров в интерфейсе
            List<string> duplicateValues = null;

            foreach (var firstEntry in interfaceParamDict)
            {
                foreach (var secondEntry in interfaceParamDict)
                {
                    if (firstEntry.Key != secondEntry.Key)
                    {
                        if (firstEntry.Value.Count >= 3 && secondEntry.Value.Count >= 3 &&
                            firstEntry.Value[1] == secondEntry.Value[1] &&
                            firstEntry.Value[2] == secondEntry.Value[2])
                        {
                            duplicateValues = new List<string> { firstEntry.Value[1], firstEntry.Value[2] };
                            break;
                        }
                    }
                }
                if (duplicateValues != null)
                {
                    break;
                }
            }

            if (duplicateValues != null)
            {
                System.Windows.Forms.MessageBox.Show("Вы пытаетесь добавить одинаковые параметры:\n" +
                    $"``{duplicateValues[0]} - {duplicateValues[1]}``\n" +
                    $"Исправьте ошибку и повторите попытку.", "Ошибка добавления параметров.",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                interfaceParamDict.Clear();
            }

            return interfaceParamDict;
        }

        // Вспомогательный метод для поиска элементов визуального дерева
        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        // Добавление параметров в семейство
        public void AddParametersToFamily(Document _doc, string activeFamilyName, Dictionary<string, List<string>> allParametersForAddDict)
        {
            string logFile = "ОТЧЁТ ОБ ОШИБКАХ.\n" + $"Параметры, которые не были добавлены в семейство {activeFamilyName}:";
            string messageBoxText = "";

            Dictionary<string, BuiltInParameterGroup> groupParameterDictionary = CreateGroupingDictionary();

            using (Transaction trans = new Transaction(_doc, "KPLN. Пакетное добавление параметров в семейство"))
            {
                trans.Start();

                foreach (var kvp in allParametersForAddDict)
                {
                    List<string> paramDetails = kvp.Value;

                    string generalParametersFileLink = paramDetails[0];
                    string parameterGroup = paramDetails[1];
                    string parameterName = paramDetails[2];
                    string typeOrInstance = paramDetails[3];

                    BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                    if (groupParameterDictionary.TryGetValue(paramDetails[4], out BuiltInParameterGroup builtInParameterGroup))
                    {
                        grouping = builtInParameterGroup;
                    }

                    revitApp.SharedParametersFilename = generalParametersFileLink;
                    DefinitionFile sharedParameterFile = revitApp.OpenSharedParameterFile();
                    DefinitionGroup parameterGroupDef = sharedParameterFile.Groups.get_Item(parameterGroup);
                    Definition parameterDefinition = parameterGroupDef.Definitions.get_Item(parameterName);

                    bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                    FamilyManager familyManager = _doc.FamilyManager;
                    ExternalDefinition externalDef = parameterDefinition as ExternalDefinition;

                    FamilyParameter existingParam = familyManager.get_Parameter(externalDef);

                    if (existingParam == null)
                    {
                        FamilyParameter familyParam = familyManager.AddParameter(externalDef, grouping, isInstance);

                        if (familyParam == null)
                        {
                            logFile += $"Error: {generalParametersFileLink}: {parameterGroup} - {parameterName}. Группирование: {paramDetails[4]} . Экземпляр: {isInstance}.\n";
                        }
                    }
                    else
                    {
                        messageBoxText += $"{parameterGroup} - {parameterName};\n";
                    }
                }

                trans.Commit();

                // Отчёт о результате выполнения в виде диалоговых окон
                if (logFile.Contains("Error"))
                {
                    System.Windows.Forms.MessageBox.Show("Параметры были добавлены добавлены в семейство с ошибками.\n" +
                        "Вы можете сохранить отчёт об ошибках в следующем диалоговом окне.", "Ошибка добавления параметров в семейство",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                    using (var saveFileDialog = new System.Windows.Forms.SaveFileDialog())
                    {
                        saveFileDialog.FileName = $"addParamLogFile_{DateTime.Now:yyyyMMddHHmmss}.txt";
                        saveFileDialog.InitialDirectory = @"X:\BIM";

                        saveFileDialog.Filter = "Отчёт об ошибках (*.txt)|*.txt";

                        if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            string filePath = saveFileDialog.FileName;

                            File.WriteAllText(filePath, logFile);
                        }
                    }
                }
                else if (messageBoxText.Contains(";"))
                {
                    System.Windows.Forms.MessageBox.Show("Все параметры добавлены в семейство, кроме тех, которые уже находятся в семействе:\n" +
                        $"{messageBoxText}", "Все параметры были добавлены в семейство",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Все параметры были добавлены в семейство", "Все параметры были добавлены в семейство",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        // Проверка наличие полей ComboBox внутри #uniqueParameterField
        public bool ExistenceCheckComboBoxes()
        {
            bool FieldsAreFilled = true;

            var uniquePanels = SP_allPanelParamsFields.Children.OfType<StackPanel>()
                        .Where(sp => sp.Tag != null && sp.Tag.ToString() == "uniqueParameterField");

            if (!uniquePanels.Any())
            {
                FieldsAreFilled = false;
            }

            return FieldsAreFilled;
        }

        // Проверка заполнености всех полей ComboBox внутри #uniqueParameterField
        public bool CheckingFillingAllComboBoxes()
        {
            bool FieldsAreFilled = true;

            var uniquePanels = SP_allPanelParamsFields.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null && sp.Tag.ToString() == "uniqueParameterField");

            foreach (var panel in uniquePanels)
            {
                var comboBoxes = panel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();

                if (comboBoxes[0].SelectedItem == null || comboBoxes[1].SelectedItem == null || comboBoxes[2].SelectedItem == null || comboBoxes[3].SelectedItem == null)
                {
                    FieldsAreFilled = false;
                }
            }

            return FieldsAreFilled;
        }

        // Изменение стиля для ComboBox "Группа" и "Параметры"
        public void makeDisabledParamGroupField()
        {
            foreach (var stackPanel in SP_allPanelParamsFields.Children.OfType<StackPanel>())
            {
                if (stackPanel.Tag?.ToString() == "uniqueParameterField")
                {
                    int count = 0;
                    foreach (var comboBox in stackPanel.Children.OfType<System.Windows.Controls.ComboBox>())
                    {
                        if (count < 2)
                        {
                            comboBox.Foreground = Brushes.Gray;
                            comboBox.IsEnabled = false;
                            count++; 
                        }
                        else
                        {
                            break; 
                        }
                    }
                }
            }
        }

        //// XAML. Открытие файла ФОПа при помощи кнопки
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckingFillingAllComboBoxes()) 
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены. Чтобы выбрать новый ФОП, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", "Не все поля заполнены", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            else
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(revitApp.SharedParametersFilename);

                if (openFileDialog.ShowDialog() == true)
                {
                    SPFPath = openFileDialog.FileName;
                    TB_filePath.Text = SPFPath;

                    HandlerGeneralParametersFile(SPFPath);
                    makeDisabledParamGroupField();
                }
            }
        }

        //// XAML. Обработчик ComboBox "Группы": обновление оригинального ComboBox "Параметры"
        private void ParamsGroupSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_paramsGroup.SelectedItem != null && CB_paramsGroup.SelectedIndex != -1)
            {
                string selectParamGroup = CB_paramsGroup.SelectedItem.ToString();
                var selectedParamName = CB_paramsName.SelectedItem;

                if (groupAndParametersFromSPFDict.ContainsKey(selectParamGroup))
                {
                    CB_paramsName.Items.Clear();                   

                    foreach (var param in groupAndParametersFromSPFDict[selectParamGroup])
                    {
                        CB_paramsName.Items.Add(param.Name);
                    }

                    if (selectedParamName != null && CB_paramsName.Items.Contains(selectedParamName))
                    {
                        CB_paramsName.SelectedItem = selectedParamName;
                    }
                }
            }    
        }

        //// XAML. Оригинальный ComboBox "Параметры": срабатывание фильтра
        private void ParamsNameTextChanged(object sender, TextChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = sender as System.Windows.Controls.ComboBox;

            if (comboBox.SelectedItem != null)
            {
                var collectionViewOriginal = CollectionViewSource.GetDefaultView(comboBox.Items);
                if (collectionViewOriginal != null)
                {
                    collectionViewOriginal.Filter = null;
                    collectionViewOriginal.Refresh();
                }
                return;
            }

            string filterText = comboBox.Text.ToLower();
            var collectionViewNew = CollectionViewSource.GetDefaultView(comboBox.Items);

            if (collectionViewNew != null)
            {
                collectionViewNew.Filter = item =>
                {
                    if (item == null) return false;
                    return item.ToString().ToLower().Contains(filterText);
                };
                collectionViewNew.Refresh();
            }
        }

        //// XAML. Оригинальный ComboBox "Параметры": уставновка фокуса и расскрытие List
        private void ParamsName_Loaded(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = sender as System.Windows.Controls.ComboBox;
            if (comboBox != null)
            {
                System.Windows.Controls.TextBox textBox = comboBox.Template.FindName("PART_EditableTextBox", comboBox) as System.Windows.Controls.TextBox;

                if (textBox != null)
                {
                    textBox.GotFocus += (s, args) => comboBox.IsDropDownOpen = true;
                }
            }
        }

        //// XAML.Оригинальный ComboBox "Параметры": обработчик открытия List
        private void ParamsNameDropDownOpened(object sender, EventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = sender as System.Windows.Controls.ComboBox;

            if (comboBox.SelectedItem != null)
            {
                comboBox.SelectedItem = null;
                comboBox.Text = string.Empty;

                var collectionViewOriginal = CollectionViewSource.GetDefaultView(comboBox.Items);

                if (collectionViewOriginal != null)
                {
                    collectionViewOriginal.Filter = null;
                    collectionViewOriginal.Refresh();
                }
            }

        }

        //// XAML.Оригинальный ComboBox "Параметры": основной обработчик событий
        private void ParamsNameSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_paramsName.SelectedItem != null)
            {
                string selectedParam = CB_paramsName.SelectedItem as String;

                foreach (var kvp in groupAndParametersFromSPFDict)
                {
                    if (kvp.Value.Any(extDef => extDef.Name == selectedParam))
                    {

                        if (CB_paramsGroup.SelectedItem == null)
                        {
                            CB_paramsGroup.SelectedItem = kvp.Key;
                        }

                        break;
                    }
                }
            }
        }

        //// XAML. Удалить оригинальный SP_panelParamFields через кнопку
        private void RemovePanel(object sender, RoutedEventArgs e)
        {
            Button buttonDel = sender as Button;
            StackPanel panel = buttonDel.Parent as StackPanel;

            if (panel != null)
            {
                System.Windows.Controls.Panel parent = panel.Parent as System.Windows.Controls.Panel;
                if (parent != null)
                {
                    parent.Children.Remove(panel);
                }
            }
        }

        //// XAML. Добавление новой панели параметров uniqueParameterField через кнопку
        private void AddPanelParamFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            StackPanel newPanel = new StackPanel
            {
                Tag = "uniqueParameterField",
                ToolTip = $"ФОП: {SPFPath}",
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(20, 0, 20, 12)
            };

            System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
            {
                Tag = SPFPath,
                ToolTip = $"ФОП: {SPFPath}",
                Width = 270,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0)
            };

            System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
            {
                Width = 490,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
            };

            foreach (var key in groupAndParametersFromSPFDict.Keys)
            {
                cbParamsGroup.Items.Add(key);
            }

            cbParamsGroup.SelectionChanged += (s, ev) =>
            {
                cbParamsName.Items.Clear();

                string selectedParamGroup = cbParamsGroup.SelectedItem.ToString();
                
                if (groupAndParametersFromSPFDict.ContainsKey(selectedParamGroup))
                {
                    foreach (var param in groupAndParametersFromSPFDict[selectedParamGroup])
                    {
                        cbParamsName.Items.Add(param.Name);
                    }
                }
            };

            System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
            {
                Width = 105,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                SelectedIndex = 0
            };

            foreach (string key in CreateTypeInstanceList())
            {
                cbTypeInstance.Items.Add(key);
            }

            System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
            {
                Width = 340,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                SelectedIndex = 21
            };

            foreach (string key in CreateGroupingDictionary().Keys)
            {
                cbGrouping.Items.Add(key);
            }

            Button removeButton = new Button
            {
                Width = 30,
                Height = 25,
                Content = "X",
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            removeButton.Click += (s, ev) =>
            {
                SP_allPanelParamsFields.Children.Remove(newPanel);
            };

            newPanel.Children.Add(cbParamsGroup);
            newPanel.Children.Add(cbParamsName);
            newPanel.Children.Add(cbTypeInstance);
            newPanel.Children.Add(cbGrouping);
            newPanel.Children.Add(removeButton);

            SP_allPanelParamsFields.Children.Add(newPanel);
        }

        //// XAML. Добавление новой панели параметров uniqueParameterField из JSON
        private void AddPanelParamFieldsJson(Dictionary<string, List<string>> allParamInInterfaceFromJsonDict)
        {
            // Уникальный ФОП для каждого uniqueParameterField параметра
            foreach (var keyDict in allParamInInterfaceFromJsonDict)
            {
                groupAndParametersFromSPFDict.Clear();

                List<string> allParamInInterfaceFromJsonValues = keyDict.Value;
                revitApp.SharedParametersFilename = allParamInInterfaceFromJsonValues[0];

                try
                {
                    DefinitionFile defFile = revitApp.OpenSharedParameterFile();

                    if (defFile == null)
                    {
                        System.Windows.Forms.MessageBox.Show($"ФОП ``{allParamInInterfaceFromJsonValues[0]}``\n" +
                        "не найден или неисправен. Работа плагина остановлена", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        break;
                    }

                    foreach (DefinitionGroup group in defFile.Groups)
                    {
                        List<ExternalDefinition> parametersList = new List<ExternalDefinition>();

                        foreach (ExternalDefinition definition in group.Definitions)
                        {
                            parametersList.Add(definition);
                        }

                        groupAndParametersFromSPFDict[group.Name] = parametersList;
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"ФОП ``{allParamInInterfaceFromJsonValues[0]}``\n" +
                        "не найден или неисправен. Работа плагина остановлена", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    break;
                }

                // Создание новой StackPanel #uniqueParameterField 
                StackPanel newPanel = new StackPanel
                {
                    Tag = "uniqueParameterField",
                    ToolTip = $"ФОП: {allParamInInterfaceFromJsonValues[0]}",
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(20, 0, 20, 12)
                };

                System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
                {
                    Tag = allParamInInterfaceFromJsonValues[0],
                    ToolTip = $"ФОП: {allParamInInterfaceFromJsonValues[0]}",
                    IsEnabled = false,
                    Width = 270,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };

                System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
                {
                    IsEnabled = false,
                    Width = 490,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };

                foreach (var key in groupAndParametersFromSPFDict.Keys)
                {
                    cbParamsGroup.Items.Add(key);
                }

                cbParamsGroup.SelectedItem = allParamInInterfaceFromJsonValues[1];

                if (!groupAndParametersFromSPFDict.ContainsKey(allParamInInterfaceFromJsonValues[1]))
                {
                    System.Windows.Forms.MessageBox.Show($"Параметр ``{allParamInInterfaceFromJsonValues[1]}`` не найден в ФОП\n" +
                        $"(``{allParamInInterfaceFromJsonValues[0]})\n" +
                        "Работа плагина остановлена.\n", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    break;
                }

                foreach (var param in groupAndParametersFromSPFDict[allParamInInterfaceFromJsonValues[1]])
                {
                    cbParamsName.Items.Add(param.Name);
                }

                cbParamsName.SelectedItem = allParamInInterfaceFromJsonValues[2];

                System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
                {
                    Width = 105,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    SelectedIndex = 0
                };

                foreach (string key in CreateTypeInstanceList())
                {
                    cbTypeInstance.Items.Add(key);
                }

                if (CreateTypeInstanceList().Contains(allParamInInterfaceFromJsonValues[3]))
                {
                    cbTypeInstance.SelectedItem = allParamInInterfaceFromJsonValues[3];
                } else 
                {
                    cbTypeInstance.SelectedIndex = -1;
                }
               
                System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
                {
                    Width = 340,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    SelectedIndex = 21
                };

                foreach (string key in CreateGroupingDictionary().Keys)
                {
                    cbGrouping.Items.Add(key);
                }

                if (CreateGroupingDictionary().ContainsKey(allParamInInterfaceFromJsonValues[4]))
                {
                    cbGrouping.SelectedItem = allParamInInterfaceFromJsonValues[4];
                }
                else
                {
                    cbGrouping.SelectedIndex = -1;
                }

                Button removeButton = new Button
                {
                    Width = 30,
                    Height = 25,
                    Content = "X",
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                    Foreground = new SolidColorBrush(Colors.White)
                };

                removeButton.Click += (s, ev) =>
                {
                    SP_allPanelParamsFields.Children.Remove(newPanel);
                };

                newPanel.Children.Add(cbParamsGroup);
                newPanel.Children.Add(cbParamsName);
                newPanel.Children.Add(cbTypeInstance);
                newPanel.Children.Add(cbGrouping);
                newPanel.Children.Add(removeButton);

                SP_allPanelParamsFields.Children.Add(newPanel);
                TB_filePath.Text = allParamInInterfaceFromJsonValues[0];
            }
        }

        //// XAML. Добавление параметров в семейство при нажатии на кнопку
        private void AddParamInFamilyButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ExistenceCheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Нет заполненных параметров.\n" +
                    "Чтобы добавить параметры в семейство, заполните хотя бы один параметр и повторите попытку.", "Нет параметров для добавления",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (!CheckingFillingAllComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены\n" +
                    "Чтобы добавить параметры в семейство, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            } else
            {
                Dictionary<string, List<string>> allParametersForAddDict = CreateInterfaceParamDict();

               AddParametersToFamily(_doc, activeFamilyName, allParametersForAddDict);
            }
        }

        //// XAML. Сохранение параметров в JSON-файл при нажатии на кнопку 
        private void SaveParamFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckingFillingAllComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены.\n" +
                    "Чтобы сохранить файл параметров, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            else if (!ExistenceCheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Нет заполненных параметров.\n" +
                    "Чтобы сохранить файл параметров, заполните хотя бы один параметр и повторите попытку.", "Нет параметров для добавления",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            else
            {
                var paramSettingsDict = CreateInterfaceParamDict();

                if (paramSettingsDict.Count == 0)
                {
                    return;
                }

                string initialDirectory = @"X:\BIM";
                string defaultFileName = $"presettingGeneralParameters_{DateTime.Now:yyyyMMddHHmmss}.json";

                using (var saveFileDialog = new System.Windows.Forms.SaveFileDialog())
                {
                    saveFileDialog.InitialDirectory = initialDirectory;
                    saveFileDialog.FileName = defaultFileName;
                    saveFileDialog.Filter = "Файл преднастроек добавления общих параметров (*.json)|*.json";
                    saveFileDialog.Title = "Сохранить файл преднастроек добавления общих параметров";

                    if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        var parameterList = new List<Dictionary<string, string>>();

                        foreach (var entry in paramSettingsDict)
                        {
                            var parameterEntry = new Dictionary<string, string>
                        {
                            { "NE", entry.Key },
                            { "pathFile", entry.Value[0] },
                            { "groupParameter", entry.Value[1] },
                            { "nameParameter", entry.Value[2] },
                            { "instance", entry.Value[3] },
                            { "grouping", entry.Value[4] }
                        };
                            parameterList.Add(parameterEntry);
                        }

                        string jsonData = JsonConvert.SerializeObject(parameterList, Formatting.Indented);

                        using (StreamWriter writer = new StreamWriter(filePath))
                        {
                            writer.Write(jsonData);
                        }

                        System.Windows.Forms.MessageBox.Show("Файл успешно сохранён по ссылке:\n" +
                            $"{filePath}", "Успех", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
        }
    }
}
