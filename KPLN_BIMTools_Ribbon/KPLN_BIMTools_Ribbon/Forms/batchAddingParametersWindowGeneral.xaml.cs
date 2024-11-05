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
        List<string> allParamNameListForSearchField = new List<string>();

        // Определение функций-словарей
        public List<string> CreateTypeInstanceList()
        {
            return batchAddingParametersWindowСhoice.CreateTypeInstanceList();
        }

        public Dictionary<string, BuiltInParameterGroup> CreateGroupingDictionary()
        {
            return batchAddingParametersWindowСhoice.CreateGroupingDictionary();
        }

        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, string parameterValue, string parameterValueDataType)
        {
            batchAddingParametersWindowСhoice choiceWindow = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);
            choiceWindow.RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParam, parameterValue, parameterValueDataType);
        }
        public string CheckingValueOfAParameter(System.Windows.Controls.ComboBox comboBox, System.Windows.Controls.TextBox textBox, string paramTypeName)
        {
            batchAddingParametersWindowСhoice choiceWindow = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);
            return choiceWindow.CheckingValueOfAParameter(comboBox, textBox, paramTypeName);
        }

#if Revit2023 || Debug2023
        public string GetParamTypeName(ExternalDefinition def, ForgeTypeId value)
        {
            batchAddingParametersWindowСhoice choiceWindow = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);
            return choiceWindow.GetParamTypeName(def, value);
        }
#endif


        public batchAddingParametersWindowGeneral(UIApplication uiapp, string activeFamilyName, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            revitApp = uiapp.Application;
            _doc = uiapp.ActiveUIDocument?.Document;

            this.activeFamilyName = activeFamilyName;
            TB_familyName.Text = activeFamilyName;

            this.jsonFileSettingPath = jsonFileSettingPath;
           
            if (!string.IsNullOrEmpty(jsonFileSettingPath))
            {
                Dictionary<string, List<string>> allParamInInterfaceFromJsonDict = ParsingDataFromJsonToInterfaceDictionary(jsonFileSettingPath);
                AddPanelParamFieldsJson(allParamInInterfaceFromJsonDict);
                jsonFileSettingPath = "";
            }
            else
            {
                FillingComboBoxTypeInstance();
                FillingComboBoxGroupingName();

                SPFPath = revitApp.SharedParametersFilename;
                HandlerGeneralParametersFile(SPFPath);
                TB_filePath.Text = SPFPath;

                SP_panelParamFields.ToolTip = $"ФОП: {SPFPath}";               
                CB_paramsGroup.ToolTip = $"ФОП: {SPFPath}";
                CB_paramsGroup.Tag = SPFPath;               
            }         
        }

        // Открытие ФОП и формирование Dictionary<String, List<ExternalDefinition>> -> List<string> параметров; 
        public void HandlerGeneralParametersFile(string filePath)
        {
            // Блокировка значений "Группа" и "Параметр" при смене ФОП
            foreach (var stackPanel in SP_allPanelParamsFields.Children.OfType<StackPanel>())
            {
                if (stackPanel.Tag?.ToString() == "uniqueParameterField")
                {
                    int count = 0;

                    foreach (var comboBox in stackPanel.Children.OfType<System.Windows.Controls.ComboBox>())
                    {
                        if (count < 2 && comboBox.SelectedItem != null)
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

            // Основной обработчик ФОП
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

                allParamNameListForSearchField = CreateAllParamNameListForSearchField(groupAndParametersFromSPFDict);

                foreach (var key in groupAndParametersFromSPFDict.Keys)
                {
                    CB_paramsGroup.Items.Add(key);
                }

                foreach (var param in allParamNameListForSearchField)
                {
                    CB_paramsName.Items.Add(param);
                }
            }
            catch (Exception)
            {
                TB_filePath.Text = "";

                System.Windows.Forms.MessageBox.Show($"{SPFPath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                return;
            }
        }

        // Создание List<string> всех параметров из ФОП для ComboBox с поиском
        static public List<string> CreateAllParamNameListForSearchField(Dictionary<String, List<ExternalDefinition>> groupAndParametersFromSPFDict)
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

        // Заполнение оригинального ComboBox "Тип/Экземпляр"
        public void FillingComboBoxTypeInstance()
        {
            foreach (string key in CreateTypeInstanceList())
            {
                CB_typeInstance.Items.Add(key);
            }

            CB_typeInstance.SelectedItem = "Тип";
        }
      
        // Заполнение оригинального ComboBox "Группирование"
        public void FillingComboBoxGroupingName()
        {
            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                CB_grouping.Items.Add(groupName);
            }

            CB_grouping.SelectedItem = "Прочее";
        }

        // Создание словаря со всеми параметрами указанными в интерфейсе
        public Dictionary<string, List<string>> CreateInterfaceParamDict()
        {
            ClearingIncorrectlyFilledFieldsParams();

            Dictionary<string, List<string>> parametersDictionary = new Dictionary<string, List<string>>();

            //// Проверка наличие хоть одного поля ComboBox внутри #uniqueParameterField
            if (!SP_allPanelParamsFields.Children.OfType<FrameworkElement>().Any(child => child.Tag?.ToString() == "uniqueParameterField"))
            {
                System.Windows.Forms.MessageBox.Show($"Нет параметров для добавления.\n"
                    + $"Исправьте это и повторите попытку.", "Параметров нет",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                return parametersDictionary;
            }

            //// Проверка на дубликаты параметров в интерфейсе
            List<string> TextBoxValues = new List<string>();
            List<string> rTextBoxValues = new List<string>();

            bool endCreateDict = false;

            var allPanels = SP_allPanelParamsFields.Children.OfType<StackPanel>().Where(panel => panel.Tag?.ToString() == "uniqueParameterField");

            foreach (var panel in allPanels)
            {
                if (panel.Children.Count > 1 && panel.Children[1] is System.Windows.Controls.ComboBox comboBox && comboBox.SelectedItem != null)
                {
                    string value = comboBox.SelectedItem.ToString();

                    if (TextBoxValues.Contains(value) && !rTextBoxValues.Contains(value))
                    {
                        rTextBoxValues.Add(value);
                    }
                    else
                    {
                        TextBoxValues.Add(value);
                    }
                }
            }

            if (rTextBoxValues.Count > 0)
            {
                string duplicatesValues = "";

                foreach (string value in rTextBoxValues)
                {
                    duplicatesValues += $"{value}\n";
                }

                System.Windows.Forms.MessageBox.Show($"Вы пытаетесь добавить одинаковые параметры:\n" +
                        $"{duplicatesValues}" +
                        $"Исправьте это и повторите попытку.", "Одинаковые параметры",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                endCreateDict = true;
            }

            //// Проверка заполнености всех полей ComboBox
            if (!CheckingFillingAllComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show($"Не все значения параметров заполнены.\n"
                    + $"Исправьте это и повторите попытку.", "Не все параметры заполнены",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                endCreateDict = true;
            }

            //// Проверка наличие правильно заполненого поля TextBox "Значение параметра"
            var uniquePanels = SP_allPanelParamsFields.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null && sp.Tag.ToString() == "uniqueParameterField");

            string gParamName = "";

            foreach (var panel in uniquePanels)
            {
                var textBox = panel.Children.OfType<System.Windows.Controls.TextBox>().FirstOrDefault();

                if (textBox != null && textBox.Tag != null && textBox.Tag.ToString() == "invalid")
                {
                    var comboBoxes = panel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();
                    gParamName += $"{comboBoxes[0].SelectedItem?.ToString()} - {comboBoxes[1].SelectedItem?.ToString()}\n";                   
                }
            }

            if (gParamName.Contains("-"))
            {
                System.Windows.Forms.MessageBox.Show($"Неправильно заполнено значение параметра:\n" +
                        $"{gParamName}" +
                        $"Исправьте это и повторите попытку.", "Неправильно заполнено значение параметра",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                endCreateDict = true;
            }

            if (endCreateDict)
            {
                return parametersDictionary;
            }

            //// Сбор данных с интерфейса для словаря
            foreach (var stackPanel in SP_allPanelParamsFields.Children.OfType<StackPanel>())
            {
                if (stackPanel.Tag != null && stackPanel.Tag.ToString() == "uniqueParameterField")
                {
                    int stackPanelIndex = parametersDictionary.Count + 1;
                    string stackPanelKey = $"ParamFromSPFAdd-{stackPanelIndex}";

                    var comboBoxes = stackPanel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();
                    var textBoxes = stackPanel.Children.OfType<System.Windows.Controls.TextBox>().ToList();

                    List<string> contentsOfDataFields = new List<string>();

                    if (comboBoxes.Count >= 4)
                    {
                        string firstComboBoxTag = comboBoxes[0].Tag?.ToString() ?? "";
                        contentsOfDataFields.Add(firstComboBoxTag);

                        foreach (var comboBox in comboBoxes.Take(4))
                        {
                            contentsOfDataFields.Add(comboBox.SelectedItem?.ToString() ?? "");
                        }

                        foreach (var textBox in textBoxes)
                        {
                            string textBoxValue = textBox.Text;

                            if (textBoxValue == "" || textBoxValue.Contains("При необходимости, вы можете указать значение параметра") 
                                || textBoxValue.Contains("Необходимо указать:") 
                                || textBoxValue.Contains("Нельзя добавить значение в данный параметр"))
                            {
                                contentsOfDataFields.Add("None");
                                contentsOfDataFields.Add(comboBoxes[1].Tag?.ToString() ?? "NoneType");
                            }
                            else
                            {
                                contentsOfDataFields.Add(textBoxValue);
                                contentsOfDataFields.Add(comboBoxes[1].Tag?.ToString() ?? "NoneType");
                            }
                        }

                        parametersDictionary.Add(stackPanelKey, contentsOfDataFields);
                    }
                }
            }

            return parametersDictionary;
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

                if (comboBoxes[0].SelectedItem == null || comboBoxes[1].SelectedItem == null
                    || comboBoxes[2].SelectedItem == null || comboBoxes[3].SelectedItem == null)
                {
                    FieldsAreFilled = false;
                }
            }

            return FieldsAreFilled;
        }

        // Очистка неправильно заполненных ComboTextBox "Параметры"
        public void ClearingIncorrectlyFilledFieldsParams()
        {
            foreach (var child in SP_allPanelParamsFields.Children)
            {
                if (child is StackPanel stackPanel && stackPanel.Tag?.ToString() == "uniqueParameterField")
                {
                    var comboBoxes = stackPanel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();

                    if (comboBoxes.Count >= 2)
                    {
                        var secondComboBox = comboBoxes[1];

                        if (secondComboBox.SelectedItem == null)
                        {
                            secondComboBox.Text = "";
                        }
                    }
                }
            }
        }

        //// XAML. Открытие файла ФОПа при помощи кнопки "Заменить"
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            ClearingIncorrectlyFilledFieldsParams();

            if (!CheckingFillingAllComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены. Чтобы выбрать новый ФОП, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", 
                    "Не все поля заполнены", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            else
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(null);

                if (openFileDialog.ShowDialog() == true)
                {
                    SPFPath = openFileDialog.FileName;
                    TB_filePath.Text = SPFPath;

                    HandlerGeneralParametersFile(SPFPath);
                }
            }
        }

        //// XAML. Поведение (!) оригинального comboBox "Группы" - Основной обработчик
        private void CB_paramsGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearingIncorrectlyFilledFieldsParams();

            if (CB_paramsGroup.SelectedItem != null)
            {
                string selectParamGroup = CB_paramsGroup.SelectedItem.ToString();
                var selectedParamName = CB_paramsName.SelectedItem;

                if (groupAndParametersFromSPFDict.ContainsKey(selectParamGroup))
                {
                    CB_paramsName.Items.Clear();
                    TB_paramValue.Text = "Выберите значение в поле ``Параметр``";

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

        //// XAML. Поведение сomboTextBox "Параметр" - Загрузка
        private void CB_paramsName_Loaded(object sender, RoutedEventArgs e)
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

        //// XAML. Поведение сomboTextBox "Параметр" - Обработчик текста
        private void CB_paramsName_TextChanged(object sender, TextChangedEventArgs e)
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

        //// XAML. Поведение сomboTextBox "Параметр" - Выпадающий список
        private void CB_paramsName_DropDownOpened(object sender, EventArgs e)
        {
            System.Windows.Controls.ComboBox comboBox = sender as System.Windows.Controls.ComboBox;

            if (comboBox.SelectedItem != null)
            {
                string comboBoxValue = comboBox.SelectedItem.ToString();
                comboBox.SelectedItem = null;
                comboBox.Text = comboBoxValue;

                var collectionViewOriginal = CollectionViewSource.GetDefaultView(comboBox.Items);

                if (collectionViewOriginal != null)
                {
                    collectionViewOriginal.Filter = null;
                    collectionViewOriginal.Refresh();
                }
            }
        }

        //// XAML. Поведение (!) оригинального сomboTextBox "Параметр" - Потеря фокуса
        private void CB_paramsName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (CB_paramsName.SelectedItem == null)
            {
                CB_paramsName.Text = "";
                TB_paramValue.IsEnabled = false;               
                TB_paramValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            }
        }

        //// XAML. Поведение (!) оригинального сomboTextBox "Параметры" - Основной обработчик 
        private void CB_paramsName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_paramsName.SelectedItem != null)
            {
                string selectedParam = CB_paramsName.SelectedItem.ToString();

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

                if (CheckingFillingAllComboBoxes())
                {
                    try
                    {
                        revitApp.SharedParametersFilename = TB_filePath.Text;

                        DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                        DefinitionGroup defGroup = defFile.Groups.get_Item(CB_paramsGroup.SelectedItem.ToString());
                        ExternalDefinition def = defGroup.Definitions.get_Item(CB_paramsName.SelectedItem.ToString()) as ExternalDefinition;

#if Revit2020 || Debug2020
                        ParameterType paramType = def.ParameterType;
#endif

#if Revit2023 || Debug2023
                        ForgeTypeId paramTypeId = def.GetDataType();
                        string paramType = GetParamTypeName(def, paramTypeId);
#endif

                        CB_paramsName.Tag = paramType;
                        TB_paramValue.IsEnabled = true;                       
                        TB_paramValue.Tag = "nonestatus";

                        if (paramType.ToString() == "Image")
                        {
                            TB_paramValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                        }
                        else if (paramType.ToString() == "LoadClassification")
                        {
                            TB_paramValue.IsEnabled = false;
                            TB_paramValue.Text = "Нельзя добавить значение в данный параметр";                            
                        }
                        else if (paramType.ToString() == "FamilyType")
                        {
                            TB_paramValue.IsEnabled = false;
                            TB_paramValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                            TB_paramValue.Tag = "invalid";
                            TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));                         
                        }
                        else 
                        {
                            TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {paramType.ToString()})";
                            TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                    }
                    catch (Exception)
                    {
                        TB_paramValue.IsEnabled = false;
                        TB_paramValue.Text = "Не удалось прочитать параметр. Тип данных: ОШИБКА";
                        TB_paramValue.Tag = "invalid";
                        TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    }
                }
            }          
        }

        //// XAML.Поведение сomboTextBox TextBox "Значение параметра" - Получение фокуса
        private void TB_paramValue_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text.ToString().Contains("При необходимости, вы можете указать значение параметра") 
                || textBox.Text.ToString().Contains("Необходимо указать:"))               
            {
                textBox.Clear();
            }
        }

        //// XAML.Поведение (!) оригинального сomboTextBox TextBox "Значение параметра" - Потеря фокуса
        private void TB_paramValue_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TB_paramValue.Text))
            {
                TB_paramValue.Tag = "nonestatus";
              
                if (CB_paramsName.Tag.ToString() == "Image")
                {
                    TB_paramValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                }
                else if (CB_paramsName.Tag.ToString() == "LoadClassification")
                {
                    TB_paramValue.IsEnabled = false;
                    TB_paramValue.Text = "Нельзя добавить значение в данный параметр";                   
                }
                else if (CB_paramsName.Tag.ToString() == "FamilyType")
                {
                    TB_paramValue.IsEnabled = false;
                    TB_paramValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                    TB_paramValue.Tag = "invalid";
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                }
                else
                {
                    TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {CB_paramsName.Tag})";
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                }                
            } 
            else 
            {
                string paramTypeName = CB_paramsName.Tag.ToString();

                if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramTypeName) == "blue" || TB_paramValue.Text.ToString().StartsWith("="))
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramTypeName) == "green")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramTypeName) == "red")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                    TB_paramValue.Tag = "invalid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramTypeName) == "yellow")
                {
                    CB_paramsName.Text = "";
                    TB_paramValue.IsEnabled = false;                                  
                    TB_paramValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                    TB_paramValue.Tag = "nonestatus";
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый    
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
            ClearingIncorrectlyFilledFieldsParams();

            StackPanel newPanel = new StackPanel
            {
                Tag = "uniqueParameterField",
                ToolTip = $"ФОП: {SPFPath}",
                Orientation = Orientation.Horizontal,
                Height = 52,
                Margin = new Thickness(20, 0, 20, 12)
            };

            System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
            {
                Tag = SPFPath,
                ToolTip = $"ФОП: {SPFPath}",
                Width = 270,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
                Padding = new Thickness(8, 4, 0, 0)
            };

            foreach (var key in groupAndParametersFromSPFDict.Keys)
            {
                cbParamsGroup.Items.Add(key);
            }

            System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
            {
                IsEditable = true,
                StaysOpenOnEdit = true,
                IsTextSearchEnabled = false,
                Width = 490,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
                Padding = new Thickness(8, 3, 0, 0),
            };

            allParamNameListForSearchField = CreateAllParamNameListForSearchField(groupAndParametersFromSPFDict);

            foreach (var param in allParamNameListForSearchField)
            {
                cbParamsName.Items.Add(param);
            }

            cbParamsName.Loaded += (s, ev) =>
            {
                if (cbParamsName != null)
                {
                    System.Windows.Controls.TextBox textBox = cbParamsName.Template.FindName("PART_EditableTextBox", cbParamsName) as System.Windows.Controls.TextBox;

                    if (textBox != null)
                    {
                        textBox.TextChanged += (os, args) =>
                        {
                            if (cbParamsName.SelectedItem != null)
                            {
                                var collectionViewOriginal = CollectionViewSource.GetDefaultView(cbParamsName.Items);
                                if (collectionViewOriginal != null)
                                {
                                    collectionViewOriginal.Filter = null;
                                    collectionViewOriginal.Refresh();
                                }
                                return;
                            }

                            string filterText = cbParamsName.Text.ToLower();
                            var collectionViewNew = CollectionViewSource.GetDefaultView(cbParamsName.Items);

                            if (collectionViewNew != null)
                            {
                                collectionViewNew.Filter = item =>
                                {
                                    if (item == null) return false;
                                    return item.ToString().ToLower().Contains(filterText);
                                };
                                collectionViewNew.Refresh();
                            }
                        };

                        textBox.GotFocus += (os, args) => cbParamsName.IsDropDownOpen = true;
                    }
                }
            };

            cbParamsName.DropDownOpened += CB_paramsName_DropDownOpened;

            System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
            {
                Width = 105,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
                Padding = new Thickness(8, 4, 0, 0),
            };

            foreach (string key in CreateTypeInstanceList())
            {
                cbTypeInstance.Items.Add(key);
            }

            cbTypeInstance.SelectedItem = "Тип";


            System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
            {
                Width = 340,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
                Padding = new Thickness(8, 4, 0, 0),
            };

            foreach (string key in CreateGroupingDictionary().Keys)
            {
                cbGrouping.Items.Add(key);
            }

            cbGrouping.SelectedItem = "Прочее";

            Button removeButton = new Button
            {
                Width = 30,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
                Content = "X",
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            removeButton.Click += RemovePanel;

            System.Windows.Controls.TextBox tbParamValue = new System.Windows.Controls.TextBox
            {
                Text = "Выберите значение в поле ``Группа`` или ``Параметр``",
                IsEnabled = false,
                Width = 1245,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(-1250, 0, 0, 0),
                Padding = new Thickness(15, 3, 0, 0),
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213))
            };

            cbParamsGroup.SelectionChanged += (s, ev) =>
            {
                ClearingIncorrectlyFilledFieldsParams();

                if (cbParamsGroup.SelectedItem != null)
                {
                    string selectParamGroup = cbParamsGroup.SelectedItem.ToString();
                    var selectedParamName = cbParamsName.SelectedItem;

                    if (groupAndParametersFromSPFDict.ContainsKey(selectParamGroup))
                    {
                        cbParamsName.Items.Clear();
                        tbParamValue.Text = "Выберите значение в поле ``Параметр``";

                        foreach (var param in groupAndParametersFromSPFDict[selectParamGroup])
                        {
                            cbParamsName.Items.Add(param.Name);
                        }

                        if (selectedParamName != null && cbParamsName.Items.Contains(selectedParamName))
                        {
                            cbParamsName.SelectedItem = selectedParamName;
                        }
                    }
                }
            };

            cbParamsName.LostFocus += (s, ev) =>
            {
                if (cbParamsName.SelectedItem == null)
                {
                    cbParamsName.Text = "";
                    tbParamValue.IsEnabled = false;
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                }
            };
       
            cbParamsName.SelectionChanged += (s, ev) =>
            {
                if (cbParamsName.SelectedItem != null)
                {
                    string selectedParam = cbParamsName.SelectedItem as String;

                    foreach (var kvp in groupAndParametersFromSPFDict)
                    {
                        if (kvp.Value.Any(extDef => extDef.Name == selectedParam))
                        {
                            if (cbParamsGroup.SelectedItem == null)
                            {
                                cbParamsGroup.SelectedItem = kvp.Key;
                            }

                            break;
                        }
                    }

                    if (CheckingFillingAllComboBoxes())
                    {
                        try
                        {
                            revitApp.SharedParametersFilename = TB_filePath.Text;

                            DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                            DefinitionGroup defGroup = defFile.Groups.get_Item(cbParamsGroup.SelectedItem.ToString());
                            ExternalDefinition def = defGroup.Definitions.get_Item(cbParamsName.SelectedItem.ToString()) as ExternalDefinition;

#if Revit2020 || Debug2020
                            ParameterType paramType = def.ParameterType;
#endif
#if Revit2023 || Debug2023
                            ForgeTypeId paramTypeId = def.GetDataType();
                            string paramType = GetParamTypeName(def, paramTypeId);
#endif

                            cbParamsName.Tag = paramType;
                            tbParamValue.IsEnabled = true;
                            tbParamValue.Tag = "nonestatus";

                            if (cbParamsName.Tag.ToString() == "Image")
                            {
                                tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                            }
                            else if (cbParamsName.Tag.ToString() == "LoadClassification")
                            {
                                tbParamValue.IsEnabled = false;
                                tbParamValue.Text = "Нельзя добавить значение в данный параметр";                                
                            }
                            else if (cbParamsName.Tag.ToString() == "FamilyType")
                            {
                                tbParamValue.IsEnabled = false;                               
                                tbParamValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                                tbParamValue.Tag = "invalid";
                                tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                            }
                            else
                            {
                                tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {paramType.ToString()})";
                                tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                            }                                                     
                        }
                        catch (Exception)
                        {
                            tbParamValue.IsEnabled = false; 
                            tbParamValue.Text = "Не удалось прочитать параметр. Тип данных: ОШИБКА";
                            cbParamsName.Tag = "invalid";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                        }
                    }
                }
            };

            tbParamValue.GotFocus += TB_paramValue_GotFocus;

            tbParamValue.LostFocus += (s, ev) =>
            {
                if (string.IsNullOrEmpty(tbParamValue.Text))
                {
                    tbParamValue.Tag = "nonestatus";

                    if (cbParamsName.Tag.ToString() == "Image")
                    {
                        tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                    }
                    else if (cbParamsName.Tag.ToString() == "LoadClassification")
                    {
                        tbParamValue.IsEnabled = false;
                        tbParamValue.Text = "Нельзя добавить значение в данный параметр";                       
                    }
                    else if (cbParamsName.Tag.ToString() == "FamilyType")
                    {
                        tbParamValue.IsEnabled = false;                       
                        tbParamValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                        tbParamValue.Tag = "invalid";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    }

                    else
                    {
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {cbParamsName.Tag})";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }                   
                }
                else
                {
                    string paramTypeName = cbParamsName.Tag.ToString(); ;

                    if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "blue" || tbParamValue.Text.ToString().StartsWith("="))
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "green")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "red")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                        tbParamValue.Tag = "invalid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "yellow")
                    {
                        cbParamsName.Text = "";
                        tbParamValue.IsEnabled = false;
                        tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                        tbParamValue.Tag = "nonestatus";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый                                                                  
                    }
                }
        };

            newPanel.Children.Add(cbParamsGroup);
            newPanel.Children.Add(cbParamsName);
            newPanel.Children.Add(cbTypeInstance);
            newPanel.Children.Add(cbGrouping);
            newPanel.Children.Add(removeButton);
            newPanel.Children.Add(tbParamValue);

            SP_allPanelParamsFields.Children.Add(newPanel);
        }

        //// XAML. Добавление параметров в семейство при нажатии на кнопку
        private void AddParamInFamilyButton_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<string>> allParametersForAddDict = CreateInterfaceParamDict();

            if (allParametersForAddDict.Count == 0)
            {
                return;
            }
            else
            {
                using (Transaction trans = new Transaction(_doc, "KPLN. Пакетное добавление параметров в семейство"))
                {
                    trans.Start();

                    FamilyManager familyManager = _doc.FamilyManager;
                    Dictionary<string, BuiltInParameterGroup> groupParameterDictionary = CreateGroupingDictionary();

                    bool successfulResult = false;
                    string problematicParametersLog = "ОТЧЁТ ОБ ОШИБКАХ.\n" + $"Параметры, которые не были добавлены в семейство {activeFamilyName}:\n";

                    foreach (var kvp in allParametersForAddDict)
                    {
                        List<string> paramDetails = kvp.Value;

                        string generalParametersFileLink = paramDetails[0];
                        string parameterGroup = paramDetails[1];
                        string parameterName = paramDetails[2];

                        string typeOrInstance = paramDetails[3];
                        bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                        string parameterValue = paramDetails[5];
                        string parameterValueDataType = paramDetails[6];

                        BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                        if (groupParameterDictionary.TryGetValue(paramDetails[4], out BuiltInParameterGroup builtInParameterGroup))
                        {
                            grouping = builtInParameterGroup;
                        }

                        revitApp.SharedParametersFilename = generalParametersFileLink;
                        DefinitionFile sharedParameterFile = revitApp.OpenSharedParameterFile();
                        DefinitionGroup parameterGroupDef = sharedParameterFile.Groups.get_Item(parameterGroup);
                        Definition parameterDefinition = parameterGroupDef.Definitions.get_Item(parameterName);
                        ExternalDefinition externalDef = parameterDefinition as ExternalDefinition;
                        FamilyParameter existingParam = familyManager.get_Parameter(externalDef);

                        try
                        {
                            if (existingParam != null)
                            {
                                familyManager.RemoveParameter(existingParam);
                            }

                            FamilyParameter familyParam = familyManager.AddParameter(externalDef, grouping, isInstance);

                            if (parameterValue.StartsWith("="))
                            {
                                string parameterFormula = parameterValue.Substring(1);
                                familyManager.SetFormula(familyParam, parameterFormula);
                            }
                            else if (parameterValue != "None")
                            {
                                RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParam, parameterValue, parameterValueDataType);
                            }

                            successfulResult = true;
                        }
                        catch (Exception ex)
                        {
                            problematicParametersLog += $"Error: {generalParametersFileLink}: {parameterGroup} - {parameterName}. Группирование: {paramDetails[4]} . Экземпляр: {isInstance}. (!) ОШИБКА ДОБАВЛЕНИЯ ФОРМУЛЫ: {parameterValue}\n";
                            paramDetails[5] = "!ОШИБКА";
                            successfulResult = false;
                        }
                    }

                    trans.Commit();

                    //// Замена элементов интерфейса если возникли ошибки (неверное значение или FamilyType)
                    if (allParametersForAddDict.Values.Any(list => list.Contains("!ОШИБКА")))
                    {
                        int index = 0;

                        foreach (StackPanel uniqueParameterField in SP_allPanelParamsFields.Children)
                        {
                            if (uniqueParameterField.Tag?.ToString() == "uniqueParameterField")
                            {
                                if (index < allParametersForAddDict.Count)
                                {
                                    List<string> parameterValues = allParametersForAddDict.ElementAt(index).Value;

                                    if (parameterValues.Count > 5)
                                    {
                                        foreach (var element in uniqueParameterField.Children)
                                        {
                                            if (element is System.Windows.Controls.TextBox textBox)
                                            {
                                                textBox.Text = parameterValues[5];

                                                if (parameterValues[5] == "!ОШИБКА")
                                                {
                                                    textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                                                    textBox.Tag = "invalid";
                                                    textBox.IsEnabled = false;
                                                }
                                                else if (parameterValues[5] == "None")
                                                {
                                                    textBox.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {parameterValues[6]})";
                                                }

                                                break;
                                            }
                                        }
                                    }

                                    index++;
                                }
                            }
                        }
                    }

                    //// Сохранение log-файла с ошибками
                    if (!successfulResult)
                    {
                        System.Windows.Forms.MessageBox.Show("Параметры были добавлены добавлены в семейство с ошибками. Вы можете сохранить отчёт об ошибках в следующем диалоговом окне.", 
                            "Ошибка добавления параметров в семейство", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                        using (var saveFileDialog = new System.Windows.Forms.SaveFileDialog())
                        {
                            saveFileDialog.FileName = $"addParamLogFile_{DateTime.Now:yyyyMMddHHmmss}.txt";
                            saveFileDialog.InitialDirectory = @"X:\BIM";

                            saveFileDialog.Filter = "Отчёт об ошибках (*.txt)|*.txt";

                            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string filePath = saveFileDialog.FileName;

                                System.IO.File.WriteAllText(filePath, problematicParametersLog);
                            }
                        }
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Все параметры были добавлены в семейство", "Все параметры были добавлены в семейство",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
        }

        //// XAML. Сохранение параметров в JSON-файл при нажатии на кнопку 
        private void SaveParamFileSettingsButton_Click(object sender, RoutedEventArgs e)
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
                                { "grouping", entry.Value[4] },
                                { "parameterValue", entry.Value[5] },
                                { "parameterValueDataType", entry.Value[6] },
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

        // JSON. Формирование словаря с параметрами из файла преднастроек
        public Dictionary<string, List<string>> ParsingDataFromJsonToInterfaceDictionary(string jsonFileSettingPath)
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
                newParamList[nameParameter].Add(entry["parameterValue"]);
                newParamList[nameParameter].Add(entry["parameterValueDataType"]);
            }

            return newParamList;           
        }

        //// JSON. Добавление новой панели параметров uniqueParameterField из файла преднастроек
        private void AddPanelParamFieldsJson(Dictionary<string, List<string>> allParamInInterfaceFromJsonDict)
        {
            StackPanel parentContainer = (StackPanel)SP_panelParamFields.Parent;
            parentContainer.Children.Remove(SP_panelParamFields);

            //// Уникальный ФОП для каждого uniqueParameterField параметра
            foreach (var keyDict in allParamInInterfaceFromJsonDict)
            {
                groupAndParametersFromSPFDict.Clear();

                List<string> allParamInInterfaceFromJsonValues = keyDict.Value;

                SPFPath = allParamInInterfaceFromJsonValues[0];
                revitApp.SharedParametersFilename = SPFPath;
                TB_filePath.Text = SPFPath;

                try
                {
                    DefinitionFile defFile = revitApp.OpenSharedParameterFile();

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
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show($"ФОП ``{SPFPath}``\n" +
                        "не найден или неисправен. Работа плагина остановлена", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    break;
                }

                //// Создание новой StackPanel #uniqueParameterField 
                StackPanel newPanel = new StackPanel
                {
                    Tag = "uniqueParameterField",
                    ToolTip = $"ФОП: {SPFPath}",
                    Orientation = Orientation.Horizontal,
                    Height = 52,
                    Margin = new Thickness(20, 0, 20, 12)
                };

                System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
                {
                    IsEnabled = false,
                    Tag = SPFPath,
                    ToolTip = $"ФОП: {SPFPath}",                    
                    Width = 270,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };

                foreach (var key in groupAndParametersFromSPFDict.Keys)
                {
                    cbParamsGroup.Items.Add(key);
                }

                System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
                {
                    IsEnabled = false,
                    Width = 490,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };

                foreach (var param in groupAndParametersFromSPFDict[allParamInInterfaceFromJsonValues[1]])
                {
                    cbParamsName.Items.Add(param.Name);
                }

                System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
                {
                    Width = 105,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                };

                foreach (string key in CreateTypeInstanceList())
                {
                    cbTypeInstance.Items.Add(key);
                }

                System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
                {
                    Width = 340,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                };

                foreach (string key in CreateGroupingDictionary().Keys)
                {
                    cbGrouping.Items.Add(key);
                }

                Button removeButton = new Button
                {
                    Width = 30,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Content = "X",
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                    Foreground = new SolidColorBrush(Colors.White)
                };

                removeButton.Click += RemovePanel;

                System.Windows.Controls.TextBox tbParamValue = new System.Windows.Controls.TextBox
                {
                    Width = 1245,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Margin = new Thickness(-1250, 0, 0, 0),
                    Padding = new Thickness(15, 3, 0, 0),
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213))
                };

                tbParamValue.Loaded += (s, ev) =>
                {
                    if (tbParamValue.Text == "None")
                    {
                        tbParamValue.Tag = "nonestatus";

                        if (allParamInInterfaceFromJsonValues[6] == "Image")
                        {
                            tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                        }
                        else if (allParamInInterfaceFromJsonValues[6] == "FamilyType")
                        {
                            tbParamValue.IsEnabled = false;                         
                            tbParamValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                            tbParamValue.Tag = "invalid";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                        }
                        else if (allParamInInterfaceFromJsonValues[6] == "LoadClassification")
                        {
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = "Нельзя добавить значение в данный параметр";                           
                        }
                        else
                        {
                            tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {allParamInInterfaceFromJsonValues[6]})";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                    }
                    else
                    {
                        string paramTypeName = allParamInInterfaceFromJsonValues[6];

                        if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "blue" || tbParamValue.Text.ToString().StartsWith("="))
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "green")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "red")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                            tbParamValue.Tag = "invalid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "yellow")
                        {
                            cbParamsName.Text = "";
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                            tbParamValue.Tag = "nonestatus";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                    }
                };

                tbParamValue.GotFocus += TB_paramValue_GotFocus;

                tbParamValue.LostFocus += (s, ev) =>
                {
                    if (string.IsNullOrEmpty(tbParamValue.Text))
                    {
                        tbParamValue.Tag = "nonestatus";

                        if (allParamInInterfaceFromJsonValues[6] == "Image")
                        {
                            tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                        }
                        else if (allParamInInterfaceFromJsonValues[6] == "LoadClassification")
                        {
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = "Нельзя добавить значение в данный параметр";                         
                        }
                        else if (allParamInInterfaceFromJsonValues[6] == "FamilyType")
                        {
                            tbParamValue.IsEnabled = false;                           
                            tbParamValue.Text = "Данный параметр добавить нельзя. Удалите данный параметр и добавьте другой";
                            tbParamValue.Tag = "invalid";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                        }
                        else
                        {
                            tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {allParamInInterfaceFromJsonValues[6]})";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                    }
                    else
                    {
                        string paramTypeName = allParamInInterfaceFromJsonValues[6];

                        if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "blue" || tbParamValue.Text.ToString().StartsWith("="))
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "green")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); 
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "red")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); 
                            tbParamValue.Tag = "invalid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramTypeName) == "yellow")
                        {
                            cbParamsName.Text = "";
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";                         
                            tbParamValue.Tag = "nonestatus";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                    }
                };

                cbParamsGroup.SelectedItem = allParamInInterfaceFromJsonValues[1];
                cbParamsName.SelectedItem = allParamInInterfaceFromJsonValues[2];
                cbParamsName.Tag = allParamInInterfaceFromJsonValues[6];
                cbTypeInstance.SelectedItem = allParamInInterfaceFromJsonValues[3];
                cbGrouping.SelectedItem = allParamInInterfaceFromJsonValues[4];
                tbParamValue.Text = allParamInInterfaceFromJsonValues[5];

                newPanel.Children.Add(cbParamsGroup);
                newPanel.Children.Add(cbParamsName);
                newPanel.Children.Add(cbTypeInstance);
                newPanel.Children.Add(cbGrouping);
                newPanel.Children.Add(removeButton);
                newPanel.Children.Add(tbParamValue);

                SP_allPanelParamsFields.Children.Add(newPanel);
            }

            if (!CheckingFillingAllComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Неисправность параметров в файле преднастроек.\n" +
                    "Проверьте файл преднастроек и повторите попытку.",
                    "Ошибка чтения ФОП", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

        }
    }
}