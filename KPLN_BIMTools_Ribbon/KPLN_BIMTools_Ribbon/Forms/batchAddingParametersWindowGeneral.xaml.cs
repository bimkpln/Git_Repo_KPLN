using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;


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

        public string generalSettingsFilePath;
        public Dictionary<String, List<ExternalDefinition>> generalParametersFileDict = new Dictionary<String, List<ExternalDefinition>>();

        public batchAddingParametersWindowGeneral(UIApplication uiapp, string activeFamilyName, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            revitApp = uiapp.Application;
            _doc = uiapp.ActiveUIDocument?.Document;

            this.activeFamilyName = activeFamilyName;
            this.jsonFileSettingPath = jsonFileSettingPath;

            TB_familyName.Text = activeFamilyName;

            ParametersOptionSelectionChanged();
            GroupingNameSelectionChanged();

            if (!string.IsNullOrEmpty(jsonFileSettingPath))
            {
                ///////////////////////////////////////////////////////////////////// ЗАГРУЗКА ИНТЕРФЕЙСА ИЗ JSON-ФАЙЛА
                generalSettingsFilePath = revitApp.SharedParametersFilename;
                TB_filePath.Text = generalSettingsFilePath;
            }
            else
            {
                generalSettingsFilePath = revitApp.SharedParametersFilename;

                HandlerGeneralParametersFile(generalSettingsFilePath);

                TB_filePath.Text = generalSettingsFilePath;
                CB_paramsGroup.Tag = generalSettingsFilePath;
            }
        }

        // Создание списка "Тип/Экземпляр"
        static public List<string> CreateParametersOptionList()
        {
            List<string> parametersOptionList = new List<string>();
            parametersOptionList.Add("Тип");
            parametersOptionList.Add("Экземпляр");

            return parametersOptionList;
        }

        // Добавление парпаметров группирования в словарь
        static public Dictionary<string, BuiltInParameterGroup> CreateGroupParameterDictionary()
        {
            Dictionary<string, BuiltInParameterGroup> groupParameters = new Dictionary<string, BuiltInParameterGroup>();

            groupParameters.Add("Аналитическая модель", BuiltInParameterGroup.PG_ANALYTICAL_MODEL);
            groupParameters.Add("Видимость", BuiltInParameterGroup.PG_VISIBILITY);
            groupParameters.Add("Второстепенный конец", BuiltInParameterGroup.PG_SECONDARY_END);
            groupParameters.Add("Выравнивание аналитической модели", BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT);
            groupParameters.Add("Геометрия разделения", BuiltInParameterGroup.PG_DIVISION_GEOMETRY);
            groupParameters.Add("Графика", BuiltInParameterGroup.PG_GRAPHICS);
            groupParameters.Add("Данные", BuiltInParameterGroup.PG_DATA);
            groupParameters.Add("Зависимости", BuiltInParameterGroup.PG_CONSTRAINTS);
            groupParameters.Add("Идентификация", BuiltInParameterGroup.PG_IDENTITY_DATA);
            groupParameters.Add("Материалы и отделка", BuiltInParameterGroup.PG_MATERIALS);
            groupParameters.Add("Механизмы", BuiltInParameterGroup.PG_MECHANICAL);
            groupParameters.Add("Механизмы - Нагрузки", BuiltInParameterGroup.PG_MECHANICAL_LOADS);
            groupParameters.Add("Механизмы - Расход", BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW);
            groupParameters.Add("Моменты", BuiltInParameterGroup.PG_MOMENTS);
            groupParameters.Add("Набор", BuiltInParameterGroup.PG_COUPLER_ARRAY);
            groupParameters.Add("Набор арматурных стержней", BuiltInParameterGroup.PG_REBAR_ARRAY);
            groupParameters.Add("Несущие конструкции", BuiltInParameterGroup.PG_STRUCTURAL);
            groupParameters.Add("Общая легенда", BuiltInParameterGroup.PG_OVERALL_LEGEND);
            groupParameters.Add("Общие", BuiltInParameterGroup.PG_GENERAL);
            groupParameters.Add("Основной конец", BuiltInParameterGroup.PG_PRIMARY_END);
            groupParameters.Add("Параметры IFC", BuiltInParameterGroup.PG_IFC);
            groupParameters.Add("Прочее", BuiltInParameterGroup.INVALID);
            groupParameters.Add("Размеры", BuiltInParameterGroup.PG_GEOMETRY);
            groupParameters.Add("Расчет несущих конструкций", BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS);
            groupParameters.Add("Расчет энергопотребления", BuiltInParameterGroup.PG_ENERGY_ANALYSIS);
            groupParameters.Add("Редактирование формы перекрытия", BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT);
            groupParameters.Add("Результат анализа", BuiltInParameterGroup.PG_ANALYSIS_RESULTS);
            groupParameters.Add("Сантехника", BuiltInParameterGroup.PG_PLUMBING);
            groupParameters.Add("Свойства модели", BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES);
            groupParameters.Add("Свойства экологически чистого здания", BuiltInParameterGroup.PG_GREEN_BUILDING);
            groupParameters.Add("Сегменты и соединительные детали", BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
            groupParameters.Add("Силы", BuiltInParameterGroup.PG_FORCES);
            groupParameters.Add("Система пожаротушения", BuiltInParameterGroup.PG_FIRE_PROTECTION);
            groupParameters.Add("Слои", BuiltInParameterGroup.PG_REBAR_SYSTEM_LAYERS);
            groupParameters.Add("Снятие связей/усилия для элемента", BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES);
            groupParameters.Add("Стадии", BuiltInParameterGroup.PG_PHASING);
            groupParameters.Add("Строительство", BuiltInParameterGroup.PG_CONSTRUCTION);
            groupParameters.Add("Текст", BuiltInParameterGroup.PG_TEXT);
            groupParameters.Add("Фотометрические", BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS);
            groupParameters.Add("Шрифт заголовков", BuiltInParameterGroup.PG_TITLE);
            groupParameters.Add("Электросети (PG_ELECTRICAL_ENGINEERING)", BuiltInParameterGroup.PG_ELECTRICAL_ENGINEERING);
            groupParameters.Add("Электросети (PG_ELECTRICAL)", BuiltInParameterGroup.PG_ELECTRICAL);
            groupParameters.Add("Электросети - Нагрузки", BuiltInParameterGroup.PG_ELECTRICAL_LOADS);
            groupParameters.Add("Электросети - Освещение", BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING);
            groupParameters.Add("Электросети - Создание цепей", BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING);

            return groupParameters;
        }

        // Обработчик ФОПа c обработчиком UpdateComboBoxField_Group(); 
        public void HandlerGeneralParametersFile(string filePath)
        {
            generalParametersFileDict.Clear();
            revitApp.SharedParametersFilename = generalSettingsFilePath;

            try
            {
                DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                if (defFile == null)
                {
                    System.Windows.Forms.MessageBox.Show($"{generalSettingsFilePath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    return;
                }

                foreach (DefinitionGroup group in defFile.Groups)
                {
                    List<ExternalDefinition> parametersList = new List<ExternalDefinition>();

                    foreach (ExternalDefinition definition in group.Definitions)
                    {
                        parametersList.Add(definition);
                    }

                    generalParametersFileDict[group.Name] = parametersList;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"{generalSettingsFilePath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                generalSettingsFilePath = "";
                TB_filePath.Text = generalSettingsFilePath;
            }

            UpdateComboBoxField_Group();
        }

        // Контент - обновление. ComboBox "Группы"
        public void UpdateComboBoxField_Group()
        {
            if (CB_paramsGroup.Items.Count == 0)
            {
                CB_paramsName.Items.Clear();
                CB_paramsGroup.Tag = generalSettingsFilePath;

                foreach (var key in generalParametersFileDict.Keys)
                {
                    CB_paramsGroup.Items.Add(key);
                }
            }
        }

        // Контент - обновление. ComboBox "Параметры"
        public void UpdateComboBoxField_Param(string selectGroup)
        {
            CB_paramsName.Items.Clear();

            if (generalParametersFileDict.ContainsKey(selectGroup))
            {
                foreach (var param in generalParametersFileDict[selectGroup])
                {
                    CB_paramsName.Items.Add(param.Name);
                }
            }
        }

        /// Контент - создание. ComboBox "Тип/Экземпляр"
        public void ParametersOptionSelectionChanged()
        {
            foreach (string key in CreateParametersOptionList())
            {
                CB_paramsType.Items.Add(key);
            }
        }

        // Контент - создание. ComboBox "Группирование"
        public void GroupingNameSelectionChanged()
        {
            foreach (string groupName in CreateGroupParameterDictionary().Keys)
            {
                CB_groupingName.Items.Add(groupName);
            }
        }

        //Проверка заполнености всех полей ComboBox
        public bool CheckComboBoxes()
        {
            bool FieldsAreFilled = true;

            var uniquePanels = SP_allPanelParamsFields.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null && sp.Tag.ToString() == "uniqueParameterField");

            foreach (var panel in uniquePanels)
            {
                var comboBoxes = panel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();

                if (comboBoxes[0].SelectedItem == null || comboBoxes[1].SelectedItem == null)
                {
                    FieldsAreFilled = false;
                }
            }

            return FieldsAreFilled;
        }

        //Проверка наличие полей ComboBox
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

        // Создание словаря со всеми параметрами указанными в интерфейсе
        public Dictionary<string, List<string>> CreateParamFileSettings()
        {
            Dictionary<string, List<string>> paramFileSettingsDictionary = new Dictionary<string, List<string>>();

            foreach (var stackPanel in FindVisualChildren<StackPanel>(this))
            {
                if (stackPanel.Tag != null && stackPanel.Tag.ToString() == "uniqueParameterField")
                {
                    int stackPanelIndex = paramFileSettingsDictionary.Count + 1;
                    string stackPanelKey = $"GeneralParamAdd-{stackPanelIndex}";

                    List<string> comboBoxValues = new List<string>();

                    var comboBoxes = stackPanel.Children.OfType<System.Windows.Controls.ComboBox>().ToList();

                    if (comboBoxes.Count >= 4)
                    {
                        string firstComboBoxTag = comboBoxes[0].Tag?.ToString() ?? "EmptyField";
                        comboBoxValues.Add(firstComboBoxTag);

                        foreach (var comboBox in comboBoxes.Take(4))
                        {
                            comboBoxValues.Add(comboBox.SelectedItem?.ToString() ?? "EmptyField");
                        }

                        paramFileSettingsDictionary.Add(stackPanelKey, comboBoxValues);
                    }
                }
            }

            return paramFileSettingsDictionary;
        }

        // Добавление параметров в семейство
        public void AddParametersToFamily(Document _doc, string activeFamilyName, Dictionary<string, List<string>> paramSettings)
        {
            Family activeFamily = GetFamilyByName(_doc, activeFamilyName);
            if (activeFamily == null)
            {
                TaskDialog.Show("Ошибка", $"Семейство {activeFamilyName} не найдено.");
                return;
            }
            else
            {
                TaskDialog.Show("Ура", $"Семейство {activeFamilyName} сейчас перед вами.");
                return;
            }

        }

        // Функция для поиска семейства по имени
        public Family GetFamilyByName(Document _doc, string familyName)
        {
            if (_doc.IsFamilyDocument)
            {
                return _doc.OwnerFamily;
            }
            return null;
        }

        // Вспомогательный метод для поиска элементов визуального дерева
        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
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

        // XAML. Открытие файла ФОПа при помощи кнопки и вызов обработчика HandlerGeneralParametersFile
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckComboBoxes()) 
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены. Чтобы выбрать новый ФОП, заполните отсутствующие данные и повторите попытку.", "Не все поля заполнены", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(revitApp.SharedParametersFilename);

                if (openFileDialog.ShowDialog() == true)
                {
                    generalSettingsFilePath = openFileDialog.FileName;
                    TB_filePath.Text = generalSettingsFilePath;

                    HandlerGeneralParametersFile(generalSettingsFilePath);
                }
            }
        }

        // XAML. Выбор значения в "Группы", для вызова обновления "Параметры"
        private void OnGroupSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_paramsGroup.SelectedItem != null)
            {
                string selectGroup = CB_paramsGroup.SelectedItem.ToString();
                UpdateComboBoxField_Param(selectGroup);
            }    
        }

        // XAML. Удалить SP_panelParamFields через кнопку
        private void RemovePanel(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            StackPanel panel = button.Parent as StackPanel;

            if (panel != null)
            {
                System.Windows.Controls.Panel parent = panel.Parent as System.Windows.Controls.Panel;
                if (parent != null)
                {
                    parent.Children.Remove(panel);
                }
            }
        }

        // XAML. Добавление новой панели параметров uniqueParameterField
        private void AddPanelParamFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            StackPanel newPanel = new StackPanel
            {
                Tag = "uniqueParameterField",
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(20, 0, 20, 12)
            };

            System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
            {
                Tag = generalSettingsFilePath,
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

            foreach (var key in generalParametersFileDict.Keys)
            {
                cbParamsGroup.Items.Add(key);
            }

            cbParamsGroup.SelectionChanged += (s, ev) =>
            {
                string selectedGroup = cbParamsGroup.SelectedItem.ToString();
                cbParamsName.Items.Clear();

                if (generalParametersFileDict.ContainsKey(selectedGroup))
                {
                    foreach (var param in generalParametersFileDict[selectedGroup])
                    {
                        cbParamsName.Items.Add(param.Name);
                    }
                }
            };

            System.Windows.Controls.ComboBox cbParamsType = new System.Windows.Controls.ComboBox
            {
                Width = 105,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                SelectedIndex = 0
            };

            foreach (string key in CreateParametersOptionList())
            {
                cbParamsType.Items.Add(key);
            }

            System.Windows.Controls.ComboBox cbGroupingName = new System.Windows.Controls.ComboBox
            {
                Width = 340,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                SelectedIndex = 21
            };

            foreach (string key in CreateGroupParameterDictionary().Keys)
            {
                cbGroupingName.Items.Add(key);
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
            newPanel.Children.Add(cbParamsType);
            newPanel.Children.Add(cbGroupingName);
            newPanel.Children.Add(removeButton);

            SP_allPanelParamsFields.Children.Add(newPanel);
        }

        // XAML. Добавление параметров в семейство
        private void AddParamInFamilyButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены\n" +
                    "Чтобы добавить параметры в семейство, заполните отсутствующие данные и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (!ExistenceCheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Нет заполненных параметров.\n" +
                    "Чтобы сохранить файл параметров, заполните хотя бы один параметр и повторите попытку.", "Нет параметров для добавления",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else
            {
                Dictionary<string, List<string>> allParametersForAddDict = CreateParamFileSettings();
                AddParametersToFamily(_doc, activeFamilyName, allParametersForAddDict);
            }
        }

        // XAML. Сохранение параметров в файл
        private void SaveParamFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Не все поля заполнены.\n" +
                    "Чтобы сохранить файл параметров, заполните отсутствующие данные и повторите попытку.", "Не все поля заполнены", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (!ExistenceCheckComboBoxes())
            {
                System.Windows.Forms.MessageBox.Show("Нет заполненных параметров.\n" +
                    "Чтобы сохранить файл параметров, заполните хотя бы один параметр и повторите попытку.", "Нет параметров для добавления",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else
            {
                var paramSettings = CreateParamFileSettings();
                string filePath = "C:\\KPLN_Debug\\test.txt";

                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    foreach (var entry in paramSettings)
                    {
                        writer.WriteLine($"{entry.Key} {{");
                        writer.WriteLine($"pathFile: {entry.Value[0]},");
                        writer.WriteLine($"groupParameter: {entry.Value[1]},");
                        writer.WriteLine($"nameParameter: {entry.Value[2]},");
                        writer.WriteLine($"optionParameter: {entry.Value[3]},");
                        writer.WriteLine($"parameterGrouping: {entry.Value[4]},");
                        writer.WriteLine("}");
                    }
                }               
                System.Windows.Forms.MessageBox.Show("Файл успешно сохранён по ссылке:\n" +
                    $"{filePath}", "Успех", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }            
    }
}
