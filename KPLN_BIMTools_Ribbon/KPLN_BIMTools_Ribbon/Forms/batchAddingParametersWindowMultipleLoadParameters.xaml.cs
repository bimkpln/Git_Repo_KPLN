using Autodesk.Revit.UI;
using System.Windows;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.IO;
using Autodesk.Revit.DB;
using System;
using System.Windows.Forms;



namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowMultipleLoadParameters : Window
    {
        UIApplication uiapp;
        public Autodesk.Revit.ApplicationServices.Application revitApp;
        public string activeFamilyName;
        string jsonFileSettingPath;

        // Определение функций-словарей
#if Revit2020 || Debug2020
        public ParameterType GetParameterTypeFromString(string dataType)
        {
            return batchAddingParametersWindowСhoice.GetParameterTypeFromString(dataType);
        }
#else
        public ForgeTypeId GetParameterTypeFromString(string dataType)
        {
            return batchAddingParametersWindowСhoice.GetParameterTypeFromString(dataType);
        }
        public ForgeTypeId GetParameterGroupFromString(string group)
        {
            return batchAddingParametersWindowСhoice.GetParameterGroupFromString(group);
        }
#endif

#if Revit2024 || Debug2024
        public Dictionary<string, ForgeTypeId> CreateGroupingDictionary()
        {
            return batchAddingParametersWindowСhoice.CreateGroupingDictionary();
        }
#else
        public Dictionary<string, BuiltInParameterGroup> CreateGroupingDictionary()
        {
            return batchAddingParametersWindowСhoice.CreateGroupingDictionary();
        }
#endif

        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, string parameterValue, string parameterValueDataType)
        {
            batchAddingParametersWindowСhoice choiceWindow = new batchAddingParametersWindowСhoice(uiapp, activeFamilyName);
            choiceWindow.RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParam, parameterValue, parameterValueDataType);
        }

        public batchAddingParametersWindowMultipleLoadParameters(UIApplication uiapp)
        {        
            InitializeComponent();
            this.uiapp = uiapp;
            revitApp = uiapp.Application;
        }

        // Разблокирование элементов интерфейса
        public void OpenInterfaceElements(string jsonFileSettingPath)
        {           
            TB_pathJson.Text = jsonFileSettingPath;
            SP_pathJsonFields.IsEnabled = false;
            SP_allPanelFamilyFields.IsEnabled = true;
            B_addParamsInFamily.IsEnabled = true;
            B_addFamilyInInterface.IsEnabled = true;
        }

        //// XAML. Открытие JSON-файла преднастроек
        private void OpenJSON_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Преднастройка (*.json)|*.json";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                jsonFileSettingPath = openFileDialog.FileName;

                string jsonContent = System.IO.File.ReadAllText(jsonFileSettingPath);
                dynamic jsonFile = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);

                if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["pathFile"] != null && item["groupParameter"] != null && item["nameParameter"] != null && item["instance"] != null
                        && item["grouping"] != null && item["parameterValue"] != null && item["parameterValueDataType"] != null))
                {
                    paramTypeStatus.Content = "Преднастройка параметров с ФОП";
                    OpenInterfaceElements(jsonFileSettingPath);

                }
                else if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["quantity"] != null && item["parameterName"] != null && item["instance"] != null && item["categoryType"] != null
                        && item["dataType"] != null && item["grouping"] != null && item["parameterValue"] != null && item["comment"] != null))
                {
                    paramTypeStatus.Content = "Преднастройка с параметрами семейства";
                    OpenInterfaceElements(jsonFileSettingPath);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Ваш JSON-файл не является файлом преднастроек или повреждён. " +
                        "Пожалуйста, выберите другой файл.", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        //// XAML. Открытие файла семейства
        private void OpenFamily_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Семейство (*.rfa)|*.rfa";
            openFileDialog.Multiselect = true;
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                foreach (string filePathChoice in openFileDialog.FileNames)
                {
                    CreateFamilyField(filePathChoice);

                    System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
                    StackPanel stackPanel = button?.Parent as StackPanel;

                    if (stackPanel != null && stackPanel.Tag?.ToString() == "uniqueFamilyField")
                    {
                        SP_allPanelFamilyFields.Children.Remove(stackPanel);
                    }
                }
            }
        }

        //// XAML. TextBox "Ссылка на файл семейства"
        private void TB_familyPath_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 255, 181));
        }

        //// XAML. Добавление ссылки на семейство в интерфейс через кнопку "Открыть" (множественный выбор)
        private void CreateFamilyField(string filePath)
        {
            StackPanel newStackPanel = new StackPanel
            {
                Tag = "uniqueFamilyField",
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Top,
                Height = 24,
                Margin = new Thickness(10, 10, 0, 0)
            };

            System.Windows.Controls.TextBox tbFamilyPath = new System.Windows.Controls.TextBox
            {
                IsReadOnly = true,
                Text = filePath,
                Width = 815,
                Padding = new Thickness(4, 3, 0, 0),
                Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFF4FBB5"))
            };

            tbFamilyPath.GotFocus += TB_familyPath_GotFocus;

            System.Windows.Controls.Button openButton = new System.Windows.Controls.Button
            {
                IsEnabled = false,
                Content = "Открыть",
                Width = 75,
                Height = 24
            };

            openButton.Click += OpenFamily_Click;

            System.Windows.Controls.Button deleteButton = new System.Windows.Controls.Button
            {
                Content = "X",
                Width = 25,
                Height = 24,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            deleteButton.Click += DeleteFamilyField_Click;

            newStackPanel.Children.Add(tbFamilyPath);
            newStackPanel.Children.Add(openButton);
            newStackPanel.Children.Add(deleteButton);

            SP_allPanelFamilyFields.Children.Add(newStackPanel);
        }

        //// XAML. Добавление ссылки на семейство в интерфейс через кнопку "Добавить семейство"
        private void B_addFamilyInInterface_Click(object sender, RoutedEventArgs e)
        {
            StackPanel newStackPanel = new StackPanel
            {
                Tag = "uniqueFamilyField",
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Top,
                Height = 24,
                Margin = new Thickness(10, 10, 0, 0)
            };

            System.Windows.Controls.TextBox tbFamilyPath = new System.Windows.Controls.TextBox
            {
                IsReadOnly = true,
                Text = "Выберите семейство",
                Width = 815,
                Padding = new Thickness(4, 3, 0, 0),
                Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#FFF4FBB5"))
            };

            tbFamilyPath.GotFocus += TB_familyPath_GotFocus;

            System.Windows.Controls.Button openButton = new System.Windows.Controls.Button
            {
                Content = "Открыть",
                Width = 75,
                Height = 24
            };

            openButton.Click += OpenFamily_Click;

            System.Windows.Controls.Button deleteButton = new System.Windows.Controls.Button
            {
                Content = "X",
                Width = 25,
                Height = 24,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            deleteButton.Click += DeleteFamilyField_Click;

            newStackPanel.Children.Add(tbFamilyPath);
            newStackPanel.Children.Add(openButton);
            newStackPanel.Children.Add(deleteButton);

            SP_allPanelFamilyFields.Children.Add(newStackPanel);
        }

        //// XAML. Удалить поле с ссылкой на семейство
        private void DeleteFamilyField_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Button buttonDel = sender as System.Windows.Controls.Button;
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

        //// XAML. Добавление параметров в семейство
        private void B_addParamsInFamily_Click(object sender, RoutedEventArgs e)
        {
            //// Поиск пустых значений в интерфейсе и их удаление
            var panelsToRemove = new List<StackPanel>();
            //// Проверка на дубликаты ссылок семейств 
            var textBoxValues = new Dictionary<string, List<System.Windows.Controls.TextBox>>();
            bool repeatValues = false;
            string duplicateValuesMessage = "";

            foreach (var child in SP_allPanelFamilyFields.Children)
            {
                if (child is StackPanel panel && panel.Tag?.ToString() == "uniqueFamilyField")
                {
                    var textBox = panel.Children.OfType<System.Windows.Controls.TextBox>().FirstOrDefault();

                    if (textBox != null && textBox.Text == "Выберите семейство")
                    {
                        panelsToRemove.Add(panel);
                    } 
                    else if (textBox != null)
                    {
                        textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 255, 181));
                        string textValue = textBox.Text;

                        if (textBoxValues.ContainsKey(textValue))
                        {
                            textBoxValues[textValue].Add(textBox);
                        }
                        else
                        {
                            textBoxValues[textValue] = new List<System.Windows.Controls.TextBox> { textBox };
                        }
                    }
                }
            }

            //// Поиск пустых значений в интерфейсе и их удаление
            foreach (var panel in panelsToRemove)
            {
                SP_allPanelFamilyFields.Children.Remove(panel);
            }

            bool hasUniqueFamilyField = SP_allPanelFamilyFields.Children.OfType<StackPanel>()
                .Any(child => child.Tag?.ToString() == "uniqueFamilyField");

            //// Проверка на дубликаты ссылок семейств 
            foreach (var pair in textBoxValues)
            {
                if (pair.Value.Count > 1)
                {
                    repeatValues = true;
                    duplicateValuesMessage += pair.Key + "\n";

                    foreach (var textBox in pair.Value)
                    {
                        textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                    }
                }
            }

            if (repeatValues)
            {
                System.Windows.Forms.MessageBox.Show(
                $"Повторяющиеся значения:\n{duplicateValuesMessage}Исправьте это и повторите попытку.", "Ошибка повторяющихся значений",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if(!hasUniqueFamilyField)
            {
                System.Windows.Forms.MessageBox.Show(
                $"Не выбрано ни одного семейства для добавления параметров из JSONB-файла параметров.", "Не выбрано ни одного семейства",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            //// Выбор дирректории для сохранения новых семейств
            else
            {
                DialogResult result = System.Windows.Forms.MessageBox.Show(
                $"Чтобы выбрать директорию для сохранения семейств с добавленными параметрами - нажмите ``OK``.\n " +
                $"Если хотите отменить действие - нажмите ``Отмена``.", "Директория для изменённых семейств",
                System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Information);

                if (result == System.Windows.Forms.DialogResult.OK)
                {

                    string selectedDirectory = string.Empty;

                    using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                    {
                        folderDialog.Description = "Выберите директорию для сохранения семейств с добавленными параметрами.";
                        folderDialog.ShowNewFolderButton = true; 

                        DialogResult resultselectedDirectory = folderDialog.ShowDialog();

                        if (!string.IsNullOrWhiteSpace(folderDialog.SelectedPath))
                        {
                            selectedDirectory = folderDialog.SelectedPath;
                        }
                        else
                        {
                            return;
                        }
                    }

                    //// Формирование Dictionary<string, List<string>> из JSON-файлов преднастроек
                    string jsonData;

                    using (StreamReader reader = new StreamReader(jsonFileSettingPath))
                    {
                        jsonData = reader.ReadToEnd();
                    }

                    var paramListFromJson = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonData);
                    var newParamList = new Dictionary<string, List<string>>();

                    //// Создание List<string> из содержимого TextBox
                    List<string> familyFieldValues = new List<string>();

                    foreach (StackPanel uniqueFieldPanel in SP_allPanelFamilyFields.Children.OfType<StackPanel>()
                        .Where(panel => panel.Tag?.ToString() == "uniqueFamilyField"))
                    {
                        System.Windows.Controls.TextBox familyPathTextBox = uniqueFieldPanel.Children.OfType<System.Windows.Controls.TextBox>().FirstOrDefault();

                        if (familyPathTextBox != null)
                        {
                            familyFieldValues.Add(familyPathTextBox.Text);
                        }
                    }

                    //// Создание Log-файла
                    string logFile = "ОТЧЁТ ОБ ОШИБКАХ.\n";
                    bool statusOperation = false;

                    ///////// JSON. Параметры с ФОП
                    if (paramTypeStatus.Content.ToString() == "Преднастройка параметров с ФОП")
                    {
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

                        foreach (string docPath in familyFieldValues)
                        {
                            try
                            {
                                Document doc = uiapp.Application.OpenDocumentFile(docPath);

                                using (Transaction trans = new Transaction(doc, "KPLN. Пакетное добавление параметров в семейство"))
                                {
                                    trans.Start();

                                    FamilyManager familyManager = doc.FamilyManager;





#if Revit2020 || Debug2020 || Revit2023 ||Debug2023
                                    Dictionary<string, BuiltInParameterGroup> groupParameterDictionary = CreateGroupingDictionary();
#else
                                    Dictionary<string, ForgeTypeId> groupParameterDictionary = CreateGroupingDictionary();
#endif




                                    foreach (var kvp in newParamList)
                                    {
                                        List<string> paramDetails = kvp.Value;

                                        string generalParametersFileLink = paramDetails[0];
                                        string parameterGroup = paramDetails[1];
                                        string parameterName = paramDetails[2];

                                        string typeOrInstance = paramDetails[3];
                                        bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                                        string parameterValue = paramDetails[5];
                                        string parameterValueDataType = paramDetails[6];

#if Revit2020 || Debug2020 || Revit2023 ||Debug2023
                                        BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                                        if (groupParameterDictionary.TryGetValue(paramDetails[4], out BuiltInParameterGroup builtInParameterGroup))
                                        {
                                            grouping = builtInParameterGroup;
                                        }
#else
                                        ForgeTypeId grouping;

                                        if (!groupParameterDictionary.TryGetValue(paramDetails[4], out grouping))
                                        {
                                            grouping = GroupTypeId.General;
                                        }
#endif

                                        revitApp.SharedParametersFilename = generalParametersFileLink;
                                        DefinitionFile sharedParameterFile = revitApp.OpenSharedParameterFile();
                                        DefinitionGroup parameterGroupDef = sharedParameterFile.Groups.get_Item(parameterGroup);
                                        Definition parameterDefinition = parameterGroupDef.Definitions.get_Item(parameterName);
                                        ExternalDefinition externalDef = parameterDefinition as ExternalDefinition;
                                        FamilyParameter existingParam = familyManager.get_Parameter(externalDef);

                                        if (existingParam != null)
                                        {
                                            if (parameterValue.StartsWith("="))
                                            {
                                                string parameterFormula = parameterValue.Substring(1);
                                                familyManager.SetFormula(existingParam, parameterFormula);
                                            }
                                            else if (parameterValue != "None")
                                            {
                                                RelationshipOfValuesWithTypesToAddToParameter(familyManager, existingParam, parameterValue, parameterValueDataType);
                                            }
                                        }
                                        else
                                        {
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
                                        }
                                    }

                                    trans.Commit();
                                }

                                string currentFileName = System.IO.Path.GetFileName(doc.PathName);
                                string newFilePath = System.IO.Path.Combine(selectedDirectory, currentFileName);

                                SaveAsOptions saveOptions = new SaveAsOptions
                                {
                                    OverwriteExistingFile = true
                                };

                                doc.SaveAs(newFilePath, saveOptions);

                                statusOperation = true;
                            }
                            catch (Exception)
                            {
                                logFile += $"ОШИБКА!: Семейство [{docPath}]: НЕ УДАЛОСЬ ОБРАБОТАТЬ.\n";
                                statusOperation = false;
                            }
                        }
                    }

                    ///////// JSON. Параметры семейства
                    else if (paramTypeStatus.Content.ToString() == "Преднастройка с параметрами семейства")
                    {
                        foreach (var entry in paramListFromJson)
                        {
                            string nameParameter = entry["NE"];

                            if (!newParamList.ContainsKey(nameParameter))
                            {
                                newParamList[nameParameter] = new List<string>();
                            }

                            newParamList[nameParameter].Add(entry["quantity"]);
                            newParamList[nameParameter].Add(entry["parameterName"]);
                            newParamList[nameParameter].Add(entry["instance"]);
                            newParamList[nameParameter].Add(entry["categoryType"]);
                            newParamList[nameParameter].Add(entry["dataType"]);
                            newParamList[nameParameter].Add(entry["grouping"]);
                            newParamList[nameParameter].Add(entry["parameterValue"]);
                            newParamList[nameParameter].Add(entry["comment"]);
                        }
                    }

                    foreach (string docPath in familyFieldValues)
                    {
                        try
                        {
                            Document doc = uiapp.Application.OpenDocumentFile(docPath);

                            using (Transaction trans = new Transaction(doc, "KPLN. Пакетное добавление параметров в семейство"))
                            {
                                trans.Start();

                                FamilyManager familyManager = doc.FamilyManager;
                                #if Revit2020 || Debug2020
                                Dictionary<string, BuiltInParameterGroup> groupParameterDictionary = CreateGroupingDictionary();
#endif

                                foreach (var kvp in newParamList)
                                {
                                    List<string> paramDetails = kvp.Value;

                                    int quantity = int.Parse(paramDetails[0]);

                                    string paramName = paramDetails[1];

                                    string typeOrInstance = paramDetails[2];
                                    bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                                    string dataType = paramDetails[4];

                                    string parameterValue = paramDetails[6];
                                    string comment = paramDetails[7];

                                    for (int i = 0; i < quantity; i++)
                                    {
                                        string fParamName = paramName + (quantity > 1 ? (i + 1).ToString() : "");

                                        FamilyParameter existingParam = familyManager.get_Parameter(fParamName);
                                        FamilyParameter familyParameter;

#if Revit2020 || Debug2020
                                        BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                                        if (groupParameterDictionary.TryGetValue(paramDetails[5], out BuiltInParameterGroup builtInParameterGroup))
                                        {
                                            grouping = builtInParameterGroup;
                                        }

                                        ParameterType paramType = GetParameterTypeFromString(dataType);

                                        if (existingParam != null)
                                        {
                                            familyParameter = existingParam;
                                        }
                                        else
                                        {
                                            familyParameter = familyManager.AddParameter(fParamName, grouping, paramType, isInstance);
                                        }
#else
                                        ForgeTypeId groupType = GetParameterGroupFromString(paramDetails[5]);
                                        ForgeTypeId paramType = GetParameterTypeFromString(dataType);

                                        if (existingParam != null)
                                        {
                                            familyParameter = existingParam;                                    
                                        }
                                        else
                                        {
                                            familyParameter = familyManager.AddParameter(fParamName, groupType, paramType, isInstance);
                                        }       
#endif
                                        if (parameterValue.StartsWith("="))
                                        {
                                            familyManager.SetFormula(familyParameter, parameterValue.Substring(1));
                                        }
                                        else if (parameterValue != "None")
                                        {
                                            RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParameter, parameterValue, dataType);
                                        }

                                        if (familyParameter != null && comment != "None")
                                        {
                                            familyManager.SetDescription(familyParameter, comment);
                                        }
                                    }
                                }

                                trans.Commit();
                            }

                            string currentFileName = System.IO.Path.GetFileName(doc.PathName);
                            string newFilePath = System.IO.Path.Combine(selectedDirectory, currentFileName);

                            SaveAsOptions saveOptions = new SaveAsOptions
                            {
                                OverwriteExistingFile = true
                            };

                            doc.SaveAs(newFilePath, saveOptions);

                            statusOperation = true;

                        }
                        catch (Exception)
                        {
                            logFile += $"ОШИБКА!: Семейство [{docPath}]: НЕ УДАЛОСЬ ОБРАБОТАТЬ.\n";
                            statusOperation = false;
                        }

                    }

                    if (statusOperation)
                    {
                        System.Windows.Forms.MessageBox.Show("Добавление параметров успешно завершено", "Успех", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Завершено с ошибками!\n" +
                            "Выберите директорию сохранения отчёта об ошибках.", "Ошибка", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        using (var saveFileDialog = new System.Windows.Forms.SaveFileDialog())
                        {
                            saveFileDialog.FileName = $"addParamLogFile_{DateTime.Now:yyyyMMddHHmmss}.txt";
                            saveFileDialog.InitialDirectory = @"X:\BIM";

                            saveFileDialog.Filter = "Отчёт об ошибках (*.txt)|*.txt";

                            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                string filePath = saveFileDialog.FileName;

                                System.IO.File.WriteAllText(filePath, logFile);
                            }
                        }
                    }

                }
            }
        }
    }
}
