using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;


namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowFamily : Window
    {
        public UIApplication uiapp;
        public Autodesk.Revit.ApplicationServices.Application revitApp;
        public Document _doc;

        public string activeFamilyName;
        public string jsonFileSettingPath;

        // Определение функций-словарей
        public List<string> CreateTypeInstanceList()
        {
            return batchAddingParametersWindowСhoice.CreateTypeInstanceList();
        }

#if Revit2020 || Debug2020
        public ParameterType GetParameterTypeFromString(string dataType)
        {
            return batchAddingParametersWindowСhoice.GetParameterTypeFromString(dataType);
        }
#endif
#if Revit2023 || Debug2023
        public ForgeTypeId GetParameterTypeFromString(string dataType)
        {
            return batchAddingParametersWindowСhoice.GetParameterTypeFromString(dataType);
        }
        public ForgeTypeId GetParameterGroupFromString(string group)
        {
            return batchAddingParametersWindowСhoice.GetParameterGroupFromString(group);
        }
#endif

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


        public batchAddingParametersWindowFamily(UIApplication uiapp, string activeFamilyName, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            revitApp = uiapp.Application;
            _doc = uiapp.ActiveUIDocument?.Document;

            this.activeFamilyName = activeFamilyName;
            this.jsonFileSettingPath = jsonFileSettingPath;

            if (!string.IsNullOrEmpty(jsonFileSettingPath))
            {
                Dictionary<string, List<string>> allInterfaceElementDictionary = ParsingDataFromJsonToInterfaceDictionary(jsonFileSettingPath);
                AddPanelParamFieldsJson(allInterfaceElementDictionary);
            }
            else
            {
                FillingComboBoxTypeInstance();
                FillingComboBoxGroupingName();
                FillingTextBoxCategoryParameterDataTypes();
            }
        }

        // Создание Dictionary<string, List<string>> для "Категории параметров" и "Тип данных"
        static public Dictionary<string, List<string>> CategoryParameterDataTypes = new Dictionary<string, List<string>>
        {
#if Revit2020 || Debug2020
            { "Общие", new List<string>(){
                "Текст", "Целое", "Число", "Длина", "Площадь", "Объем (Общие)", "Угол", "Уклон (Общие)", "Денежная единица", "Массовая плотность", "Время", "Скорость (Общие)", "URL", "Материал", 
                "Изображение", "Да/Нет", "Многострочный текст"} },
            { "Несущие конструкции", new List<string>(){
                "Усилие", "Распределенная нагрузка по линии", "Распределенная нагрузка", "Момент", "Линейный момент", "Напряжение", "Удельный вес", "Вес", "Масса (Несущие конструкции)", "Масса на единицу площади", 
                "Коэффициент теплового расширения", "Сосредоточенный коэффициент упругости", "Линейный коэффициент упругости", "Коэффициент упругости среды", "Сосредоточенный угловой коэффициент упругости", 
                "Линейный угловой коэффициент упругости", "Смещение/прогиб", "Вращение", "Период", "Частота (Несущие конструкции)", "Пульсация", "Скорость (Несущие конструкции)", "Ускорение", "Энергия (Несущие конструкции)", 
                "Объем арматуры", "Длина армирования", "Армирование по площади", "Армирование по площади на единицу длины", "Интервал арматирования", "Защитный слой арматирования", "Диаметр стержня", "Ширина трещины", 
                "Размеры сечения", "Свойство сечения", "Площадь сечения", "Момент сопротивления сечения", "Момент инерции", "Постоянная перекоса", "Масса на единицу длины (Несущие конструкции)", "Вес на единицу длины", 
                "Площадь поверхности на единицу длины"}},
            { "ОВК", new List<string>(){
                "Плотность (ОВК)", "Трение (ОВК)", "Мощность (ОВК)", "Удельная мощность (ОВК)", "Давление (ОВК)", "Температура (ОВК)", "Разность температур (ОВК)", "Скорость (ОВК)", "Воздушный поток", "Размер воздуховода", 
                "Поперечный разрез", "Теплоприток", "Шероховатость (ОВК)", "Динамическая вязкость (ОВК)", "Плотность воздушного потока", "Холодильная нагрузка", "Отопительная нагрузка", "Холодильная нагрузка на единицу площади", 
                "Отопительная нагрузка на единицу площади", "Холодильная нагрузка на единицу объема", "Отопительная нагрузка на единицу объема", "Воздушный поток на единицу объема", "Воздушный поток, отнесенный к холодильной нагрузке", 
                "Площадь, отнесенная к холодильной нагрузке", "Площадь на единицу отопительной нагрузки", "Уклон (ОВК)", "Коэффициент", "Толщина изоляции воздуховода", "Толщина внутренней изоляции воздуховода"}},
            { "Электросети", new List<string>(){
                "Ток", "Электрический потенциал", "Частота (Электросети)", "Освещенность", "Яркость", "Световой поток", "Сила света", "Эффективность", "Мощность (ElectricalWattage)", "Мощность (ElectricalPower)", "Цветовая температура", 
                "Полная установленная мощность", "Удельная мощность (Электросети)", "Электрическое удельное сопротивление", "Диаметр провода", "Температура (Электросети)", "Разность температур (Электросети)", "Размер кабельного лотка", 
                "Размер короба", "Коэффициент спроса нагрузки", "Количество полюсов", "Классификация нагрузок"}},
            { "Трубопроводы", new List<string>(){
                "Плотность (Трубопроводы)", "Расход", "Трение (Трубопроводы)", "Давление (Трубопроводы)", "Температура (Трубопроводы)", "Разность температур (Трубопроводы)", "Скорость (Трубопроводы)", 
                "Динамическая вязкость (Трубопроводы)", "Размер трубы", "Шероховатость (Трубопроводы)", "Объем (Трубопроводы)", "Уклон (Трубопроводы)", "Толщина изоляции трубы", "Размер трубы (PipeSize)", "Размер трубы (PipeDimension)", 
                "Масса (Трубопроводы)", "Масса на единицу длины (Трубопроводы)", "Расход приборов"}},
            { "Энергия", new List<string>(){
                "Энергия (Энергия)", "Коэффициент теплопередачи", "Термостойкость", "Тепловая нагрузка", "Теплопроводность", "Удельная теплоемкость", "Удельная теплоемкость парообразования", "Проницаемость"}}
#endif
#if Revit2023 || Debug2023
            { "Общие", new List<string>(){
                "Текст", "Целое", "Угол", "Площадь", "Стоимость на единицу площади", "Расстояние", "Длина", "Массовая плотность", "Число", "Угол поворота", "Уклон (Общие)", "Скорость (Общие)", "Время", "Объем (Общие)", 
                "Денежная единица", "URL", "Материал", "Образец заливки", "Изображение", "Да/Нет", "Многострочный текст"}},
            { "Электросети", new List<string>() {
                "Полная установленная мощность", "Полная удельная мощность", "Размер кабельного лотка", "Цветовая температура", "Размер короба", "Норма затрат на электроэнергию", "Норма затрат на энергопотребление", "Ток", 
                "Коэффициент спроса нагрузки", "Эффективность", "Частота (Электросети)", "Освещенность", "Яркость", "Световой поток", "Сила света", "Электрический потенциал", "Мощность (Электросети)",
                "Удельная мощность (Электросети)", "Мощность на единицу длины", "Электрическое удельное сопротивление", "Температура (Электросети)", "Перепад температур (Электросети)", 
                "Активная мощность", "Диаметр провода", "Количество полюсов","Классификация нагрузок"}},
            { "Энергия", new List<string>() {
                "Энергия (Энергия)", "Нагревающая способность на единицу площади", "Коэффициент теплопередачи", "Изотермическая влагоемкость", "Проницаемость", "Удельная теплоемкость", "Удельная теплоемкость парообразования", 
                "Теплопроводность", "Коэффициент температурного градиента для влагоемкости", "Тепловая нагрузка", "Термостойкость"}},
            { "ОВК", new List<string>() {
                "Воздушный поток", "Плотность воздушного потока", "Воздушный поток, разделенный на холодильную нагрузку", "Воздушный поток, разделенный на объем", "Угловая скорость", "Площадь, отнесенная к холодильной нагрузке", 
                "Площадь на единицу отопительной нагрузки", "Холодильная нагрузка", "Холодильная нагрузка на единицу площади", "Холодильная нагрузка, разделенная на объем", "Поперечное сечение", "Плотность (ОВК)", 
                "Коэффициент температуропроводимости", "Толщина изоляции воздуховода", "Толщина внутренней изоляции воздуховода", "Размер воздуховода", "Коэффициент", "Поток на единицу мощности", "Трение (ОВК)", 
                "Теплоприток","Отопительная нагрузка", "Отопительная нагрузка на единицу площади", "Отопительная нагрузка, разделенная на объем", "Масса на единицу времени (ОВК)", "Мощность (ОВК)", 
                "Удельная мощность (ОВК)", "Мощность на единицу потока", "Давление (ОВК)", "Шероховатость (ОВК)", "Уклон (ОВК)", "Температура (ОВК)", "Перепад температур (ОВК)", "Скорость (ОВК)", 
                "Динамическая вязкость (ОВК)"}},
            { "Инфраструктура", new List<string>() {
                "Пикетаж", "Интенрвал пикетов"}},
            { "Трубопроводы", new List<string>() {
                "Плотность (Трубопровод)", "Расход", "Трение (Трубопроводы)", "Масса (Трубопроводы)", "Масса на единицу времени (Трубопроводы)", "Величина трубы", "Толщина изоляции трубы", "Масса на единицу длины (Трубопроводы)", 
                "Размер трубы", "Давление (Трубопроводы)", "Шероховатость (Трубопроводы)", "Уклон (Трубопроводы)", "Температура (Трубопроводы)", "Перепад температур (Трубопроводы)", "Скорость (Трубопроводы)", 
                "Динамическая вязкость (Трубопроводы)", "Объем (Трубопроводы)"}},
            { "Несущие конструкции", new List<string>() {
                "Ускорение", "Распределенная нагрузка", "Коэффициент упругости среды", "Диаметр стержня", "Ширина трещины", "Смещение/прогиб", "Энергия (Несущие конструкции)", "Усилие", "Частота (Несущие конструкции)", 
                "Линейный коэффициент упругости", "Распределенная нагрузка по линии", "Линейный момент", "Масса (Несущие конструкции)", "Масса на единицу площади", "Масса на единицу длины (Несущие конструкции)", "Момент", 
                "Момент инерции", "Период", "Сосредоточенный коэффициент упругости", "Пульсация", "Площадь армирования","Армирование по площади на единицу длины", "Защитный слой армирования", "Длина армирования",
                "Интервал армирования", "Объем арматуры", "Поворот", "Линейный угловой коэффициент упругости", "Сосредоточенный угловой коэффициент упругости", "Площадь сечения", "Размеры сечения", "Момент сопротивления сечения", 
                "Свойство сечения", "Напряжение", "Площадь поверхности на единицу длины", "Коэффициент теплового расширения", "Удельный вес", "Скорость (Несущие конструкции)", "Постоянная перекоса", "Вес", "Вес на единицу длины"}}
#endif
        };
 
        // Регулярное выражение для целых чисел (для кол-во параметров)
        public static bool IsTextAllowed(string text)
        {
            return Regex.IsMatch(text, @"^[0-9]+$");
        }

        // Заполнение оригинального ComboBox "Тип/Экземпляр"
        public void FillingComboBoxTypeInstance()
        {
            foreach (string key in CreateTypeInstanceList())
            {
                CB_typeInstance.Items.Add(key);
            }
            CB_typeInstance.SelectedIndex = 0;
        }

        // Заполнение оригинального ComboBox "Категории параметров"
        public void FillingTextBoxCategoryParameterDataTypes()
        {
            foreach (string categoryDataType in CategoryParameterDataTypes.Keys)
            {
                CB_categoryDataType.Items.Add(categoryDataType);
            }

            CB_categoryDataType.SelectedItem = "Общие";
        }

        // Заполнение оригинального ComboBox "Группирование"
        public void FillingComboBoxGroupingName()
        {
            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                CB_grouping.Items.Add(groupName);
            }
            CB_grouping.SelectedItem = "Размеры";
        }

        // Создание словаря со всеми параметрами указанными в интерфейсе
        public Dictionary<string, List<string>> CreateInterfaceParamDict()
        {
            var parametersDictionary = new Dictionary<string, List<string>>();

            //// Проверка на неверные значения параметров в интерфейсе
            bool uniqueParameterFieldFound = false;
            string errorValueString = "";

            foreach (var child in SP_allPanelParamsFields.Children)
            {
                if (child is StackPanel panel && panel.Tag?.ToString() == "uniqueParameterField")
                {
                    uniqueParameterFieldFound = true;

                    foreach (var innerChild in panel.Children)
                    {
                        if (innerChild is System.Windows.Controls.TextBox textBox && textBox.Tag?.ToString() == "invalid")
                        {
                            textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                            errorValueString += $"{textBox.Text.ToString()};\n";
                        }
                    }
                }
            }

            if (!uniqueParameterFieldFound)
            {
                System.Windows.Forms.MessageBox.Show($"Нет параметров для добавления.\n"
                    + $"Исправьте это и повторите попытку.", "Параметров нет",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                return parametersDictionary;
            }
            else if (errorValueString.Contains("Введите имя параметра"))
            {
                System.Windows.Forms.MessageBox.Show($"Не все имена параметров заполнены.\n"
                    + $"Исправьте это и повторите попытку.", "Имена не заполнены",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                return parametersDictionary;
            }
            else if (errorValueString.Contains(";"))
            {
                System.Windows.Forms.MessageBox.Show($"Значения параметров с ошибками:\n{errorValueString}"
                    + $"Исправьте ошибки и повторите попытку.", "Ошибка добавления параметров",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                return parametersDictionary;
            }

            //// Проверка на дубликаты параметров в интерфейсе
            Dictionary<string, List<System.Windows.Controls.TextBox>> textBoxValues = new Dictionary<string, List<System.Windows.Controls.TextBox>>();

            foreach (var child in SP_allPanelParamsFields.Children)
            {
                if (child is StackPanel panel && panel.Tag?.ToString() == "uniqueParameterField")
                {
                    if (panel.Children[1] is System.Windows.Controls.TextBox textBox)
                    {
                        string value = textBox.Text;

                        if (!string.IsNullOrEmpty(value))
                        {
                            if (!textBoxValues.ContainsKey(value))
                            {
                                textBoxValues[value] = new List<System.Windows.Controls.TextBox>();
                            }
                            textBoxValues[value].Add(textBox);
                        }
                    }
                }
            }

            List<string> duplicateValues = new List<string>();
            foreach (var entry in textBoxValues)
            {
                if (entry.Value.Count > 1)
                {
                    duplicateValues.Add(entry.Key);
                    foreach (var textBox in entry.Value)
                    {
                        textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                    }
                }
            }

            if (duplicateValues.Any())
            {
                string dublicateValueString = "";

                foreach (var textBox in duplicateValues)
                {
                    dublicateValueString += $"{textBox};\n";
                }

                System.Windows.Forms.MessageBox.Show($"Параметры c одинаковыми именами:\n{dublicateValueString}"
                    + $"Исправьте ошибку и повторите попытку.", "Ошибка добавления параметров.",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                return parametersDictionary;
            }

            //// Сбор данных с интерфейса для словаря
            int count = 1;

            var mainStackPanel = SP_allPanelParamsFields;

            foreach (var child in mainStackPanel.Children)
            {
                if (child is StackPanel panel && (string)panel.Tag == "uniqueParameterField")
                {
                    var values = new List<string>();

                    foreach (var element in panel.Children)
                    {
                        if (element is System.Windows.Controls.TextBox textBox)
                        {
                            if (textBox.Text.ToString().StartsWith("При необходимости, вы можете указать") 
                                || textBox.Text.ToString() == "Вы не можете заполнить значение для данного параметра")
                            {
                                values.Add ("None");
                            }
                            else 
                            {
                                values.Add(textBox.Text);
                            }
                        }
                        else if (element is System.Windows.Controls.ComboBox comboBox)
                        {
                            values.Add(comboBox.SelectedItem?.ToString() ?? string.Empty);
                        }
                    }

                    parametersDictionary.Add($"customParam-{count}", values);
                    count++;
                }
            }
         
            return parametersDictionary;
        }
     
        //// XAML. Поведение textBox "Кол-во" - Ввод текста
        private void TB_quantity_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        //// XAML. Поведение textBox "Кол-во" - Потеря фокуса
        private void TB_quantity_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (int.TryParse(textBox.Text, out int value))
            {
                if (value < 1)
                    textBox.Text = "1";
                else if (value > 100)
                    textBox.Text = "100";
            }
            else
            {
                TB_quantity.Text = "1";
            }
        }

        //// XAML. Поведение textBox "Кол-во" - Прекращение ввода текста
        private void TB_quantity_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
                e.Handled = true;
        }

        //// XAML. Поведение textBox "Имя параметра" - Получение фокуса
        private void TB_paramsName_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "Введите имя параметра" || textBox.Text == "Используется недопустимый символ")
            {
                textBox.Text = "";
                textBox.Tag = "invalid";
            }
        }

        //// XAML. Поведение textBox "Имя параметра" - Потеря фокуса
        private void TB_paramsName_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "")
            {
                textBox.Text = "Введите имя параметра";
                textBox.Tag = "invalid";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
            }
            else if (textBox.Text.ToString().Contains(";") 
                || textBox.Text.ToString().Contains("{") || textBox.Text.ToString().Contains("}")
                || textBox.Text.ToString().Contains("[") || textBox.Text.ToString().Contains("]"))
            {
                textBox.Text = "Используется недопустимый символ";
                textBox.Tag = "invalid";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
            }
            else
            {
                textBox.Tag = "valid";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));               
            }
        }

        //// XAML. Поведение (!) оригинальный comboBox "Категория" - Основной обработчик
        private void CB_categoryDataType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CB_dataType.Items.Clear();

            string selectcategoryDataType = CB_categoryDataType.SelectedItem.ToString();

            foreach (var dataType in CategoryParameterDataTypes[selectcategoryDataType])
            {
                CB_dataType.Items.Add(dataType);
            }

            if (CB_categoryDataType.SelectedItem.ToString() == "Общие")
            {
                CB_dataType.SelectedItem = "Длина";
            }
            else
            {
                CB_dataType.SelectedIndex = 0;
            }
        }

        //// XAML. Поведение (!) оригинальный comboBox "Тип данных" - Основной обработчик
        private void CB_dataType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_dataType.SelectedItem != null) 
            {
                CB_dataType.ToolTip = CB_dataType.SelectedItem.ToString();

                if (CB_dataType.SelectedItem.ToString() == "Классификация нагрузок")
                {
                    TB_paramValue.IsEnabled = false;
                    TB_paramValue.Text = "Вы не можете заполнить значение для данного параметра";
                }
                else if (CB_dataType.SelectedItem.ToString() == "Изображение")
                {
                    TB_paramValue.IsEnabled = true;
                    TB_paramValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                }
                else
                {
                    TB_paramValue.IsEnabled = true;
                    TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
                }

                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            }
        }

        //// XAML. Поведение textBox "Значение параметра" - Получение фокуса
        private void TB_paramValue_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text.Contains("При необходимости, вы можете указать значение параметра") 
                || textBox.Text.ToString().StartsWith("Необходимо указать:"))
            {
                textBox.Text = "";
            }
        }

        //// XAML. Поведение (!) оригинальный textBox "Значение параметра" - Загрузка
        private void TB_paramValue_Loaded(object sender, RoutedEventArgs e)
        {
            TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
        }

        //// XAML. Поведение (!) оригинальный textBox "Значение параметра" - Потеря фокуса
        private void TB_paramValue_LostFocus(object sender, RoutedEventArgs e)
        {
            if (CB_dataType.SelectedItem.ToString() == "Изображение" && TB_paramValue.Text == "")
            {
                TB_paramValue.IsEnabled = true;
                TB_paramValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
            }
            else if (TB_paramValue.Text == "")
            {
                TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
                TB_paramValue.Tag = "invalid";
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            }
            else
            {
                if (CheckingValueOfAParameter(CB_dataType, TB_paramValue, CB_dataType.SelectedItem.ToString()) == "blue" || TB_paramValue.Text.ToString().StartsWith("="))
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_dataType, TB_paramValue, CB_dataType.SelectedItem.ToString()) == "green")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_dataType, TB_paramValue, CB_dataType.SelectedItem.ToString()) == "red")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    TB_paramValue.Tag = "invalid";
                }
            }
        }

        //// XAML. Поведение textBox "Подсказка" - Получение фокуса
        private void TB_comment_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "При необходимости, вы можете указать описание подсказки")
            {
                textBox.Text = "";
            }
        }

        //// XAML. Поведение textBox "Подсказка" - Потеря фокуса
        private void TB_comment_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "")
            {
                textBox.Text = "При необходимости, вы можете указать описание подсказки";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            }
            else
            {
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
            }
        }

        //// XAML. Удалить SP_panelParamFields через кнопку
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
                Height = 90,
                Margin = new Thickness(10, 8, 0, 0),
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Top
            };

            System.Windows.Controls.TextBox tbQuantity = new System.Windows.Controls.TextBox
            {
                Text = "1",
                Width = 45,
                TextWrapping = TextWrapping.Wrap,
                Height = 25,
                Padding = new Thickness(5, 3, 0, 0),
                VerticalAlignment = VerticalAlignment.Top,
            };

            tbQuantity.PreviewTextInput += TB_quantity_PreviewTextInput;
            tbQuantity.PreviewKeyDown += TB_quantity_PreviewKeyDown;
            tbQuantity.LostFocus += TB_quantity_LostFocus;
            
            System.Windows.Controls.TextBox tbParamsName = new System.Windows.Controls.TextBox
            {
                Tag = "invalid",
                Text = "Введите имя параметра",
                TextWrapping = TextWrapping.Wrap,
                Width = 450,
                Height = 25,
                Padding = new Thickness(5, 3, 0, 0),
                VerticalAlignment = VerticalAlignment.Top,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)),
            };

            tbParamsName.GotFocus += TB_paramsName_GotFocus;
            tbParamsName.LostFocus += TB_paramsName_LostFocus;
            
            System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
            {
                SelectedIndex = 0,
                Width = 105,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (string key in CreateTypeInstanceList())
            {
                cbTypeInstance.Items.Add(key);
            }

            System.Windows.Controls.ComboBox cbCategoryDataType = new System.Windows.Controls.ComboBox
            {
                Width = 180,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (string categoryDataType in CategoryParameterDataTypes.Keys)
            {
                cbCategoryDataType.Items.Add(categoryDataType);
            }

            cbCategoryDataType.SelectedIndex = 0;

            System.Windows.Controls.ComboBox cbDataType = new System.Windows.Controls.ComboBox
            {
                Width = 240,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (var dataType in CategoryParameterDataTypes["Общие"])
            {
                cbDataType.Items.Add(dataType);
            }

            cbDataType.SelectedItem = "Длина";

            cbCategoryDataType.SelectionChanged += (s, ev) =>
            {
                cbDataType.Items.Clear();

                string selectcategoryDataType = cbCategoryDataType.SelectedItem.ToString();

                foreach (var dataType in CategoryParameterDataTypes[selectcategoryDataType])
                {
                    cbDataType.Items.Add(dataType);
                }

                if (cbCategoryDataType.SelectedItem.ToString() == "Общие")
                {
                    cbDataType.SelectedItem = "Длина";
                }
                else
                {
                    cbDataType.SelectedIndex = 0;
                }
            };

            System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
            {
                Width = 310,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                cbGrouping.Items.Add(groupName);
            }
            cbGrouping.SelectedItem = "Размеры";

            Button btnRemove = new Button
            {
                Content = "X",
                Width = 30,
                Height = 25,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Top
            };

            btnRemove.Click += RemovePanel;

            System.Windows.Controls.TextBox tbParamValue = new System.Windows.Controls.TextBox
            {
                Width = 1360,
                Height = 25,
                Margin = new Thickness(-1360, 0, 0, 40),
                Padding = new Thickness(5, 3, 0, 0),
                VerticalAlignment = VerticalAlignment.Bottom,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)),
                TextWrapping = TextWrapping.Wrap
            };

            tbParamValue.GotFocus += TB_paramValue_GotFocus;

            tbParamValue.LostFocus += (s, ev) =>
            {

                if (cbDataType.SelectedItem.ToString() == "Изображение" && tbParamValue.Text == "")
                {
                    tbParamValue.IsEnabled = true;
                    tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                }
                else if (tbParamValue.Text == "")
                {
                    tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                    tbParamValue.Tag = "invalid";
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                }
                else
                {
                    if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "blue" || tbParamValue.Text.ToString().StartsWith("="))
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); 
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "green")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); 
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "red")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                        tbParamValue.Tag = "invalid";
                    }
                }
            };

            tbParamValue.Loaded += (s, ev) =>
            {
                tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
            };

            cbDataType.SelectionChanged += (s, ev) =>
            {
                if (cbDataType.SelectedItem != null)
                {
                    cbDataType.ToolTip = cbDataType.SelectedItem.ToString();

                    if (cbDataType.SelectedItem.ToString() == "Классификация нагрузок")
                    {
                        tbParamValue.IsEnabled = false;
                        tbParamValue.Text = "Вы не можете заполнить значение для данного параметра";
                    }
                    else if (cbDataType.SelectedItem.ToString() == "Изображение")
                    {
                        tbParamValue.IsEnabled = true;
                        tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); 
                    }
                    else
                    {
                        tbParamValue.IsEnabled = true;
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                    }

                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                }
            };

            System.Windows.Controls.TextBox tbComment = new System.Windows.Controls.TextBox
            {
                Text = "При необходимости, вы можете указать описание подсказки",
                Width = 1360,
                Height = 40,
                Margin = new Thickness(-1360, 0, 0, 0),
                Padding = new Thickness(5, 3, 0, 0),
                VerticalAlignment = VerticalAlignment.Bottom,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)),
                TextWrapping = TextWrapping.Wrap
            };

            tbComment.GotFocus += TB_comment_GotFocus;
            tbComment.LostFocus += TB_comment_LostFocus;

            newPanel.Children.Add(tbQuantity);
            newPanel.Children.Add(tbParamsName);
            newPanel.Children.Add(cbTypeInstance);
            newPanel.Children.Add(cbCategoryDataType);
            newPanel.Children.Add(cbDataType);
            newPanel.Children.Add(cbGrouping);
            newPanel.Children.Add(btnRemove);
            newPanel.Children.Add(tbParamValue);
            newPanel.Children.Add(tbComment);

            SP_allPanelParamsFields.Children.Add(newPanel);
        }

        //// XAML. Добавление параметров в семейство при нажатии на кнопку 
        private void AddParamsInFamilyButton_Click(object sender, RoutedEventArgs e)
        {
            var allParamSettingsDict = CreateInterfaceParamDict();

            if (allParamSettingsDict.Count == 0)
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
                    string problematicParametersLog = "ОТЧЁТ ОБ ОШИБКАХ.\n" + $"Параметры, которые не были добавлены в семейство {activeFamilyName}:\n\n";

                    foreach (var kvp in allParamSettingsDict)
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

                            if (existingParam != null)
                            {
                                familyManager.RemoveParameter(existingParam);
                            }

                            try
                            {

#if Revit2020 || Debug2020
                                BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                                if (groupParameterDictionary.TryGetValue(paramDetails[5], out BuiltInParameterGroup builtInParameterGroup))
                                {
                                    grouping = builtInParameterGroup;
                                }

                                ParameterType paramType = GetParameterTypeFromString(dataType);
                                FamilyParameter familyParameter = familyManager.AddParameter(fParamName, grouping, paramType, isInstance);
#endif
#if Revit2023 || Debug2023
                                ForgeTypeId groupType = GetParameterGroupFromString(paramDetails[5]);
                                ForgeTypeId paramType = GetParameterTypeFromString(dataType);

                                FamilyParameter familyParameter = familyManager.AddParameter(fParamName, groupType, paramType, isInstance);
#endif

                                if (parameterValue.StartsWith("="))
                                {
                                    familyManager.SetFormula(familyParameter, parameterValue.Substring(1));
                                }
                                else if (parameterValue != "None")
                                {
                                    RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParameter, parameterValue, dataType);
                                }

                                if (familyParameter != null)
                                {
                                    familyManager.SetDescription(familyParameter, comment);
                                }

                                successfulResult = true;
                            }
                            catch 
                            {
                                problematicParametersLog += $"Error: КОЛ-ВО: {quantity}. ИМЯ_ПАРАМЕТРА: {paramName}. ЭКЗЕМПЛЯР: {isInstance}. ТИП ДАННЫХ: {dataType}. ГРУППИРОВАНИЕ: {paramDetails[5]}.\n" +
                                    $"ОПИСАНИЕ ПОДСКАЗКИ: {comment}\n\n";

                                successfulResult = false;
                            }
                        }                       
                    }

                    trans.Commit();

                    if (successfulResult)
                    {
                        System.Windows.Forms.MessageBox.Show("Все параметры добавлены в семейство", "Успех", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    else
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

                                System.IO.File.WriteAllText(filePath, problematicParametersLog);
                            }
                        }
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
            string defaultFileName = $"customParameters_{DateTime.Now:yyyyMMddHHmmss}.json";

            using (var saveFileDialog = new System.Windows.Forms.SaveFileDialog())
            {
                saveFileDialog.InitialDirectory = initialDirectory;
                saveFileDialog.FileName = defaultFileName;
                saveFileDialog.Filter = "Файл преднастроек добавления параметров семейства (*.json)|*.json";
                saveFileDialog.Title = "Сохранить файл преднастроек добавления параметров семейства";

                if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    var parameterList = new List<Dictionary<string, string>>();
                   
                    foreach (var entry in paramSettingsDict)
                    {
                        var parameterEntry = new Dictionary<string, string>
                        {
                            { $"NE", entry.Key },
                            { "quantity", entry.Value[0] },
                            { "parameterName", entry.Value[1] },
                            { "instance", entry.Value[2] },
                            { "categoryType", entry.Value[3] },
                            { "dataType", entry.Value[4] },
                            { "grouping", entry.Value[5] },
                            { "parameterValue", entry.Value[6] },
                            { "comment", entry.Value[7] },
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

                newParamList[nameParameter].Add(entry["quantity"]);
                newParamList[nameParameter].Add(entry["parameterName"]);
                newParamList[nameParameter].Add(entry["instance"]);
                newParamList[nameParameter].Add(entry["categoryType"]);
                newParamList[nameParameter].Add(entry["dataType"]);
                newParamList[nameParameter].Add(entry["grouping"]);
                newParamList[nameParameter].Add(entry["parameterValue"]);
                newParamList[nameParameter].Add(entry["comment"]);
            }

            return newParamList;
        }

        //// JSON. Добавление новой панели параметров uniqueParameterField из файла преднастроек
        private void AddPanelParamFieldsJson(Dictionary<string, List<string>> allParamInInterfaceFromJsonDict)
        {
            StackPanel stackPanel = this.FindName("SP_panelParamFields") as StackPanel;
            bool notSupportedVersion = false;

            if (stackPanel != null)
            {
                var parent = stackPanel.Parent as System.Windows.Controls.Panel;
                if (parent != null)
                {
                    parent.Children.Remove(stackPanel);
                }
            }

            foreach (var keyDict in allParamInInterfaceFromJsonDict)
            {
                List<string> allParamInInterfaceFromJsonList = keyDict.Value;

                StackPanel newPanel = new StackPanel
                {
                    Tag = "uniqueParameterField",
                    Height = 90,
                    Margin = new Thickness(10, 8, 0, 0),
                    Orientation = Orientation.Horizontal,
                    VerticalAlignment = VerticalAlignment.Top
                };

                System.Windows.Controls.TextBox tbQuantity = new System.Windows.Controls.TextBox
                {
                    Width = 45,
                    TextWrapping = TextWrapping.Wrap,
                    Height = 25,
                    Padding = new Thickness(5, 3, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top,
                };

                tbQuantity.PreviewTextInput += TB_quantity_PreviewTextInput;
                tbQuantity.PreviewKeyDown += TB_quantity_PreviewKeyDown;
                tbQuantity.LostFocus += TB_quantity_LostFocus;

                System.Windows.Controls.TextBox tbParamsName = new System.Windows.Controls.TextBox
                {
                    Tag = "valid",
                    TextWrapping = TextWrapping.Wrap,
                    Width = 450,
                    Height = 25,
                    Padding = new Thickness(5, 3, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)),
                };

                tbParamsName.GotFocus += TB_paramsName_GotFocus;
                tbParamsName.LostFocus += TB_paramsName_LostFocus;

                System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
                {
                    Width = 105,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                foreach (string key in CreateTypeInstanceList())
                {
                    cbTypeInstance.Items.Add(key);
                }
              
                System.Windows.Controls.ComboBox cbCategoryDataType = new System.Windows.Controls.ComboBox
                {
                    Width = 180,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                foreach (string categoryDataType in CategoryParameterDataTypes.Keys)
                {
                    cbCategoryDataType.Items.Add(categoryDataType);
                }

                System.Windows.Controls.ComboBox cbDataType = new System.Windows.Controls.ComboBox
                {
                    Width = 240,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                if (CategoryParameterDataTypes.ContainsKey(allParamInInterfaceFromJsonList[3]))
                {
                    foreach (var dataType in CategoryParameterDataTypes[allParamInInterfaceFromJsonList[3]])
                    {
                        cbDataType.Items.Add(dataType);
                    }
                }

                System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
                {
                    Width = 310,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                foreach (string groupName in CreateGroupingDictionary().Keys)
                {
                    cbGrouping.Items.Add(groupName);
                }

                Button btnRemove = new Button
                {
                    Content = "X",
                    Width = 30,
                    Height = 25,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                    Foreground = Brushes.White,
                    VerticalAlignment = VerticalAlignment.Top
                };

                btnRemove.Click += RemovePanel;

                System.Windows.Controls.TextBox tbParamValue = new System.Windows.Controls.TextBox
                {
                    Width = 1360,
                    Height = 25,
                    Margin = new Thickness(-1360, 0, 0, 40),
                    Padding = new Thickness(5, 3, 0, 0),
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)),
                    TextWrapping = TextWrapping.Wrap
                };

                cbCategoryDataType.SelectionChanged += (s, ev) =>
                {
                    cbDataType.Items.Clear();

                    string selectcategoryDataType = cbCategoryDataType.SelectedItem.ToString();

                    foreach (var dataType in CategoryParameterDataTypes[selectcategoryDataType])
                    {
                        cbDataType.Items.Add(dataType);
                    }

                    if (cbCategoryDataType.SelectedItem.ToString() == "Общие")
                    {
                        cbDataType.SelectedItem = "Длина";
                    }
                    else
                    {
                        cbDataType.SelectedIndex = 0;                      
                    }
                };

                tbParamValue.GotFocus += TB_paramValue_GotFocus;

                tbParamValue.LostFocus += (s, ev) =>
                {
                    if (cbDataType.SelectedItem.ToString() == "Изображение" && tbParamValue.Text == "")
                    {
                        tbParamValue.IsEnabled = true;
                        tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                    }
                    else if (tbParamValue.Text == "")
                    {
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                        tbParamValue.Tag = "invalid";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                    else
                    {
                        if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "blue" || tbParamValue.Text.ToString().StartsWith("="))
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "green")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbDataType, tbParamValue, cbDataType.SelectedItem.ToString()) == "red")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                            tbParamValue.Tag = "invalid";
                        }

                    }
                };

                cbDataType.SelectionChanged += (s, ev) =>
                {
                    if (cbDataType.SelectedItem != null)
                    {
                        cbDataType.ToolTip = cbDataType.SelectedItem.ToString();

                        if (cbDataType.SelectedItem.ToString() == "Классификация нагрузок")
                        {
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = "Вы не можете заполнить значение для данного параметра";
                        }
                        else if (cbDataType.SelectedItem.ToString() == "Изображение")
                        {
                            tbParamValue.IsEnabled = true;
                            tbParamValue.Text = "При необходимости, вы можете указать значение параметра в формате - [Буква диска]:\\Имя_папки\\изображение.расширение";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180));
                        }
                        else
                        {
                            tbParamValue.IsEnabled = true;
                            tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                        }

                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                };

                System.Windows.Controls.TextBox tbComment = new System.Windows.Controls.TextBox
                {
                    Width = 1360,
                    Height = 40,
                    Margin = new Thickness(-1360, 0, 0, 0),
                    Padding = new Thickness(5, 3, 0, 0),
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)),
                    TextWrapping = TextWrapping.Wrap
                };

                tbComment.GotFocus += TB_comment_GotFocus;
                tbComment.LostFocus += TB_comment_LostFocus;

                tbQuantity.Text = allParamInInterfaceFromJsonList[0];
                tbParamsName.Text = allParamInInterfaceFromJsonList[1];
                cbTypeInstance.SelectedItem = allParamInInterfaceFromJsonList[2];
                cbCategoryDataType.SelectedItem = allParamInInterfaceFromJsonList[3];
                cbDataType.SelectedItem = allParamInInterfaceFromJsonList[4];
                cbGrouping.SelectedItem = allParamInInterfaceFromJsonList[5];

                if (cbDataType.SelectedItem == null)
                {
                    tbParamValue.Text = "Данный тип данных не поддерживается текущей версией Revit";
                }
                else if (allParamInInterfaceFromJsonList[6] == "None")
                {
                    tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                }
                else
                {
                    tbParamValue.Text = allParamInInterfaceFromJsonList[6];
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                }

                if (allParamInInterfaceFromJsonList[7] == "None")
                {
                    tbComment.Text = "При необходимости, вы можете указать описание подсказки";
                }
                else
                {
                    tbComment.Text = allParamInInterfaceFromJsonList[7];
                    tbComment.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                }
              
                newPanel.Children.Add(tbQuantity);
                newPanel.Children.Add(tbParamsName);
                newPanel.Children.Add(cbTypeInstance);
                newPanel.Children.Add(cbCategoryDataType);
                newPanel.Children.Add(cbDataType);
                newPanel.Children.Add(cbGrouping);
                newPanel.Children.Add(btnRemove);
                newPanel.Children.Add(tbParamValue);
                newPanel.Children.Add(tbComment);

                SP_allPanelParamsFields.Children.Add(newPanel);

                if (cbCategoryDataType.SelectedItem == null || cbDataType.SelectedItem == null || cbGrouping.SelectedItem == null)
                {
                    notSupportedVersion = true;

                    tbQuantity.IsEnabled = false;
                    tbQuantity.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    tbParamsName.IsEnabled = false;
                    tbParamsName.Tag = "invalid";
                    tbParamsName.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    cbTypeInstance.IsEnabled = false;
                    cbCategoryDataType.IsEnabled = false;
                    cbDataType.IsEnabled = false;
                    cbGrouping.IsEnabled = false;
                    tbParamValue.IsEnabled = false;
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    tbComment.IsEnabled = false;
                    tbComment.Text = "ОШИБКА: Преднастройка для данного параметра использовала атрибуты недоступные в данной версии Revit.\n" +
                        "Чтобы продолжить - вы должны удалить данный параметр.";
                    tbComment.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                }
            }

            if (notSupportedVersion) 
            {
                System.Windows.Forms.MessageBox.Show("Данная версия Revit не поддерживает некоторые параметры из выбранного файла конфигураций.\n" +
                    "Поля с параметрами, неподдерживаемыми данной версией Revit, остались незаполненые.\n" +
                    "Крайне не рекомендуется продолжать работу плагина с этим файлом конфигураций.", 
                    "Ошибка версии Revit", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }
    }
}