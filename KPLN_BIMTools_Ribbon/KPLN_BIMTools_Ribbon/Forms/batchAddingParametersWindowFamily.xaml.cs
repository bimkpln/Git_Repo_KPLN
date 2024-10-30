using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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

        public Dictionary<string, BuiltInParameterGroup> CreateGroupingDictionary()
        {
            return batchAddingParametersWindowСhoice.CreateGroupingDictionary();
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

        // Словарь Dictionary<string, List<string>> для "Категории параметров" и "Тип данных"
        static public Dictionary<string, List<string>> CategoryParameterDataTypes = new Dictionary<string, List<string>>
        {
            { "Общие", new List<string>(){
                "Текст", "Целое", "Число", "Длина", "Площадь", "Объем (Общие)", "Угол (Общие)", "Уклон", "Денежная единица", "Массовая плотность", "Время", "Скорость (Общие)", "URL", "Материал", "Изображение", "Да/Нет", "Многострочный текст"} },
            { "Несущие конструкции", new List<string>(){
                "Усилие", "Распределенная нагрузка по линии", "Распределенная нагрузка", "Момент", "Линейный момент", "Напряжение", "Удельный вес", "Вес", "Масса (Несущие конструкции)", "Масса на единицу площади", "Коэффициент теплового расширения",
                "Сосредоточенный коэффициент упругости", "Линейный коэффициент упругости", "Коэффициент упругости среды", "Сосредоточенный угловой коэффициент упругости", "Линейный угловой коэффициент упругости", "Смещение/прогиб",
                "Вращение", "Период", "Частота (Несущие конструкции)", "Пульсация", "Скорость (Несущие конструкции)", "Ускорение", "Энергия (Несущие конструкции)", "Объем арматуры", "Длина армирования", "Армирование по площади", "Армирование по площади на единицу длины", "Интервал арматирования",
                "Защитный слой арматирования", "Диаметр стержня", "Ширина трещины", "Размеры сечения", "Свойство сечения", "Площадь сечения", "Момент сопротивления сечения", "Момент инерции", "Постоянная перекоса",
                "Масса на единицу длины (Несущие конструкции)", "Вес на единицу длины", "Площадь поверхности на единицу длины"} },
            { "ОВК", new List<string>(){
                "Плотность (ОВК)", "Трение (ОВК)", "Мощность", "Удельная мощность (ОВК)", "Давление (ОВК)", "Температура (ОВК)", "Разность температур (ОВК)", "Скорость (ОВК)", "Воздушный поток", "Размер воздуховода", "Поперечный разрез", "Теплоприток", "Шероховатость (ОВК)",
                "Динамическая вязкость (ОВК)", "Плотность воздушного потока", "Холодильная нагрузка", "Отопительная нагрузка", "Холодильная нагрузка на единицу площади", "Отопительная нагрузка на единицу площади", "Холодильная нагрузка на единицу объема",
                "Отопительная нагрузка на единицу объема", "Воздушный поток на единицу объема", "Воздушный поток, отнесенный к холодильной нагрузке", "Площадь, отнесенная к холодильной нагрузке", "Площадь на единицу отопительной нагрузки", "Уклон (ОВК)",
                "Коэффициент", "Толщина изоляции воздуховода", "Толщина внутренней изоляции воздуховода" } },
            { "Электросети", new List<string>(){
                "Ток", "Электрический потенциал", "Частота (Электросети)", "Освещенность", "Яркость", "Световой поток", "Сила света", "Эффективность", "Мощность (ElectricalWattage)", "Мощность (ElectricalPower)", "Цветовая температура", "Полная установленная мощность",
                "Удельная мощность (Электросети)", "Электрическое удельное сопротивление", "Диаметр провода", "Температура (Электросети)", "Разность температур (Электросети)", "Размер кабельного лотка", "Размер короба", "Коэффициент спроса нагрузки", "Количество полюсов", "Классификация нагрузок" } },
            { "Трубопроводы", new List<string>(){
                "Плотность (Трубопроводы)", "Расход", "Трение (Трубопроводы)", "Давление (Трубопроводы)", "Температура (Трубопроводы)", "Разность температур (Трубопроводы)", "Скорость (Трубопроводы)", "Динамическая вязкость (Трубопроводы)", "Размер трубы", "Шероховатость (Трубопроводы)",
                "Объем (Трубопроводы)", "Уклон", "Толщина изоляции трубы", "Размер трубы (PipeSize)", "Размер трубы (PipeDimension)", "Масса (Трубопроводы)", "Масса на единицу длины (Трубопроводы)", "Расход приборов" } },
            { "Энергия", new List<string>(){
                "Энергия (Энергия)", "Коэффициент теплопередачи", "Термостойкость", "Тепловая нагрузка", "Теплопроводность", "Удельная теплоемкость", "Удельная теплоемкость парообразования", "Проницаемость" } }
        };

        // Регулярное выражение для целых чисел (для кол-во параметров)
        public static bool IsTextAllowed(string text)
        {
            return Regex.IsMatch(text, @"^[0-9]+$");
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
            CB_typeInstance.SelectedIndex = 0;
        }

        // Заполнение оригинального ComboBox "Категории параметров"
        public void FillingTextBoxCategoryParameterDataTypes()
        {
            foreach (string categoryDataType in CategoryParameterDataTypes.Keys)
            {
                CB_categoryDataType.Items.Add(categoryDataType);
            }

            CB_categoryDataType.SelectedIndex = 0;
        }

        // Заполнение оригинального ComboBox "Группирование"
        public void FillingComboBoxGroupingName()
        {
            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                CB_grouping.Items.Add(groupName);
            }
            CB_grouping.SelectedIndex = 22;
        }

        // Создание словаря со всеми параметрами указанными в интерфейсе
        public Dictionary<string, List<string>> CreateInterfaceParamDict()
        {
            var parametersDictionary = new Dictionary<string, List<string>>();

            //// Проверка на неверные значения параметров в интерфейсе
            string errorValueString = "";

            foreach (var child in SP_allPanelParamsFields.Children)
            {
                if (child is StackPanel panel && panel.Tag?.ToString() == "uniqueParameterField")
                {
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

            if (errorValueString.Contains("Введите имя параметра"))
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
                            if (textBox.Text.ToString().StartsWith("При необходимости, вы можете указать"))
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

        //// XAML. Поведение оригинальный textBox "Кол-во" - Ввод текста
        private void TB_quantity_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        //// XAML. Поведение оригинальный textBox "Кол-во" - Потеря фокуса
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

        //// XAML. Поведение оригинальный textBox "Кол-во" - Прекращение ввода текста
        private void TB_quantity_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
                e.Handled = true;
        }


        //// XAML. Поведение оригинальный textBox "Имя параметра" - Получение фокуса
        private void TB_paramsName_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "Введите имя параметра")
            {
                textBox.Text = "";
                textBox.Tag = "invalid";
            }
        }

        //// XAML. Поведение оригинальный textBox "Имя параметра" - Потеря фокуса
        private void TB_paramsName_LostFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "")
            {
                textBox.Text = "Введите имя параметра";
                textBox.Tag = "invalid";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
            }
            else
            {
                textBox.Tag = "valid";
                textBox.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));               
            }
        }

        //// XAML. Поведение оригинальный comboBox "Категория" - Основной обработчик
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
                CB_dataType.SelectedIndex = 3;
            }
            else
            {
                CB_dataType.SelectedIndex = 0;
            }
        }

        //// XAML. Поведение оригинальный comboBox "Тип данных" - Основной обработчик
        private void CB_dataType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_dataType.SelectedItem != null) 
            {
                CB_dataType.ToolTip = CB_dataType.SelectedItem.ToString();
                TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
            }
        }

        //// XAML. Поведение оригинальный textBox "Значение параметра" - Загрузка
        private void TB_paramValue_Loaded(object sender, RoutedEventArgs e)
        {
            TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
        }

        //// XAML. Поведение оригинальный textBox "Значение параметра" - Получение фокуса
        private void TB_paramValue_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text.Contains("При необходимости, вы можете указать значение параметра"))
            {
                textBox.Text = "";
            }
        }

        //// XAML. Поведение оригинальный textBox "Значение параметра" - Потеря фокуса
        private void TB_paramValue_LostFocus(object sender, RoutedEventArgs e)
        {
            if (TB_paramValue.Text == "")
            {
                TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра [{CB_dataType.SelectedItem.ToString()}]";
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            }
            else
            {
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
            }
        }

        //// XAML. Поведение оригинальный textBox "Подсказка" - Получение фокуса
        private void TB_comment_GotFocus(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;

            if (textBox.Text == "При необходимости, вы можете указать описание подсказки")
            {
                textBox.Text = "";
            }
        }

        //// XAML. Поведение оригинальный textBox "Подсказка" - Потеря фокуса
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
                Width = 210,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (var dataType in CategoryParameterDataTypes["Общие"])
            {
                cbDataType.Items.Add(dataType);
            }

            cbDataType.SelectedIndex = 3;

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
                    cbDataType.SelectedIndex = 3;
                }
                else
                {
                    cbDataType.SelectedIndex = 0;
                }
            };

            System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
            {
                Width = 340,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (string groupName in CreateGroupingDictionary().Keys)
            {
                cbGrouping.Items.Add(groupName);
            }
            cbGrouping.SelectedIndex = 22;

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
                if (tbParamValue.Text == "")
                {
                    tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                }
                else
                {
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
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
                    tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
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
                System.Windows.Forms.MessageBox.Show("Для добавления параметров исправьте все ошибки", 
                    "Предупреждение", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Asterisk);
            }
            else
            {
                using (Transaction trans = new Transaction(_doc, "KPLN. Пакетное добавление параметров в семейство"))
                {
                    trans.Start();

                    FamilyManager familyManager = _doc.FamilyManager;
                    Dictionary<string, BuiltInParameterGroup> groupParameterDictionary = CreateGroupingDictionary();

                    foreach (var kvp in allParamSettingsDict)
                    {                       
                        List<string> paramDetails = kvp.Value;

                        int quantity = int.Parse(paramDetails[0]);

                        string paramName = paramDetails[1];
                        
                        string typeOrInstance = paramDetails[2];
                        bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                        string dataType = paramDetails[4]; 

                        BuiltInParameterGroup grouping = BuiltInParameterGroup.INVALID;

                        if (groupParameterDictionary.TryGetValue(paramDetails[5], out BuiltInParameterGroup builtInParameterGroup))
                        {
                            grouping = builtInParameterGroup;
                        }

                        string paramValue = paramDetails[6];
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
                                var paramType = GetParameterTypeFromString(dataType);
                                FamilyParameter familyParameter = familyManager.AddParameter(fParamName, grouping, paramType, isInstance);

                                if (familyParameter != null)
                                {
                                    familyManager.SetDescription(familyParameter, comment);
                                }
                            }
                            catch { }
                        }
                    }

                    trans.Commit();
                }
            }
        }





#if Revit2020 || Debug2020
        public ParameterType GetParameterTypeFromString(string dataType)
        {
            switch (dataType)
            {
                case "Текст": return ParameterType.Text;
                case "Целое": return ParameterType.Integer;
                case "Число": return ParameterType.Number;
                case "Длина": return ParameterType.Length;
                case "Площадь": return ParameterType.Area;
                case "Объем (Общие)": return ParameterType.Volume;
                case "Угол": return ParameterType.Angle;
                case "Уклон (Общие)": return ParameterType.Slope;
                case "Денежная единица": return ParameterType.Currency;
                case "Массовая плотность": return ParameterType.MassDensity;
                case "Время": return ParameterType.TimeInterval;
                case "Скорость (Общие)": return ParameterType.Speed;
                case "URL": return ParameterType.URL;
                case "Материал": return ParameterType.Material;
                case "Изображение": return ParameterType.Image;
                case "Да/Нет": return ParameterType.YesNo;
                case "Многострочный текст": return ParameterType.MultilineText;
                case "Усилие": return ParameterType.Force;
                case "Распределенная нагрузка по линии": return ParameterType.LinearForce;
                case "Распределенная нагрузка": return ParameterType.AreaForce;
                case "Момент": return ParameterType.Moment;
                case "Линейный момент": return ParameterType.LinearMoment;
                case "Напряжение": return ParameterType.Stress;
                case "Удельный вес": return ParameterType.UnitWeight;
                case "Вес": return ParameterType.Weight;
                case "Масса (Несущие конструкции)": return ParameterType.Mass;
                case "Масса на единицу площади": return ParameterType.MassPerUnitArea;
                case "Коэффициент теплового расширения": return ParameterType.ThermalExpansion;
                case "Сосредоточенный коэффициент упругости": return ParameterType.ForcePerLength;
                case "Линейный коэффициент упругости": return ParameterType.LinearForcePerLength;
                case "Коэффициент упругости среды": return ParameterType.AreaForcePerLength;
                case "Сосредоточенный угловой коэффициент упругости": return ParameterType.ForceLengthPerAngle;
                case "Линейный угловой коэффициент упругости": return ParameterType.LinearForceLengthPerAngle;
                case "Смещение/прогиб": return ParameterType.DisplacementDeflection;
                case "Вращение": return ParameterType.Rotation;
                case "Период": return ParameterType.Period;
                case "Частота (Несущие конструкции)": return ParameterType.StructuralFrequency;
                case "Пульсация": return ParameterType.Pulsation;
                case "Скорость (Несущие конструкции)": return ParameterType.StructuralVelocity;
                case "Ускорение": return ParameterType.Acceleration;
                case "Энергия (Несущие конструкции)": return ParameterType.Energy;
                case "Объем арматуры": return ParameterType.ReinforcementVolume;
                case "Длина армирования": return ParameterType.ReinforcementLength;
                case "Армирование по площади": return ParameterType.ReinforcementArea;
                case "Армирование по площади на единицу длины": return ParameterType.ReinforcementAreaPerUnitLength;
                case "Интервал арматирования": return ParameterType.ReinforcementSpacing;
                case "Защитный слой арматирования": return ParameterType.ReinforcementCover;
                case "Диаметр стержня": return ParameterType.BarDiameter;
                case "Ширина трещины": return ParameterType.CrackWidth;
                case "Размеры сечения": return ParameterType.SectionDimension;
                case "Свойство сечения": return ParameterType.SectionProperty;
                case "Площадь сечения": return ParameterType.SectionArea;
                case "Момент сопротивления сечения": return ParameterType.SectionModulus;
                case "Момент инерции": return ParameterType.MomentOfInertia;
                case "Постоянная перекоса": return ParameterType.WarpingConstant;
                case "Масса на единицу длины (Несущие конструкции)": return ParameterType.MassPerUnitLength;
                case "Вес на единицу длины": return ParameterType.WeightPerUnitLength;
                case "Площадь поверхности на единицу длины": return ParameterType.SurfaceArea;
                case "Плотность (ОВК)": return ParameterType.HVACDensity;
                case "Трение (ОВК)": return ParameterType.HVACFriction;
                case "Мощность": return ParameterType.HVACPower;
                case "Удельная мощность (ОВК)": return ParameterType.HVACPowerDensity;
                case "Давление (ОВК)": return ParameterType.HVACPressure;
                case "Температура (ОВК)": return ParameterType.HVACTemperature;
                case "Разность температур (ОВК)": return ParameterType.HVACTemperatureDifference;
                case "Скорость (ОВК)": return ParameterType.HVACVelocity;
                case "Воздушный поток": return ParameterType.HVACAirflow;
                case "Размер воздуховода": return ParameterType.HVACDuctSize;
                case "Поперечный разрез": return ParameterType.HVACCrossSection;
                case "Теплоприток": return ParameterType.HVACHeatGain;
                case "Шероховатость (ОВК)": return ParameterType.HVACRoughness;
                case "Динамическая вязкость (ОВК)": return ParameterType.HVACViscosity;
                case "Плотность воздушного потока": return ParameterType.HVACAirflowDensity;
                case "Холодильная нагрузка": return ParameterType.HVACCoolingLoad;
                case "Отопительная нагрузка": return ParameterType.HVACHeatingLoad;
                case "Холодильная нагрузка на единицу площади": return ParameterType.HVACCoolingLoadDividedByArea;
                case "Отопительная нагрузка на единицу площади": return ParameterType.HVACHeatingLoadDividedByArea;
                case "Холодильная нагрузка на единицу объема": return ParameterType.HVACCoolingLoadDividedByVolume;
                case "Отопительная нагрузка на единицу объема": return ParameterType.HVACHeatingLoadDividedByVolume;
                case "Воздушный поток на единицу объема": return ParameterType.HVACAirflowDividedByVolume;
                case "Воздушный поток, отнесенный к холодильной нагрузке": return ParameterType.HVACAirflowDividedByCoolingLoad;
                case "Площадь, отнесенная к холодильной нагрузке": return ParameterType.HVACAreaDividedByCoolingLoad;
                case "Площадь на единицу отопительной нагрузки": return ParameterType.HVACAreaDividedByHeatingLoad;
                case "Уклон (ОВК)": return ParameterType.HVACSlope;
                case "Коэффициент": return ParameterType.HVACFactor;
                case "Толщина изоляции воздуховода": return ParameterType.HVACDuctInsulationThickness;
                case "Толщина внутренней изоляции воздуховода": return ParameterType.HVACDuctLiningThickness;
                case "Ток": return ParameterType.ElectricalCurrent;
                case "Электрический потенциал": return ParameterType.ElectricalPotential;
                case "Частота (Электросети)": return ParameterType.ElectricalFrequency;
                case "Освещенность": return ParameterType.ElectricalIlluminance;
                case "Яркость": return ParameterType.ElectricalLuminance;
                case "Световой поток": return ParameterType.ElectricalLuminousFlux;
                case "Сила света": return ParameterType.ElectricalLuminousIntensity;
                case "Эффективность": return ParameterType.ElectricalEfficacy;
                case "Мощность (ElectricalWattage)": return ParameterType.ElectricalWattage;
                case "Мощность (ElectricalPower)": return ParameterType.ElectricalPower;
                case "Цветовая температура": return ParameterType.ColorTemperature;
                case "Полная установленная мощность": return ParameterType.ElectricalApparentPower;
                case "Удельная мощность (Электросети)": return ParameterType.ElectricalPowerDensity;
                case "Электрическое удельное сопротивление": return ParameterType.ElectricalResistivity;
                case "Диаметр провода": return ParameterType.WireSize;
                case "Температура (Электросети)": return ParameterType.ElectricalTemperature;
                case "Разность температур (Электросети)": return ParameterType.ElectricalTemperatureDifference;
                case "Размер кабельного лотка": return ParameterType.ElectricalCableTraySize;
                case "Размер короба": return ParameterType.ElectricalConduitSize; 
                case "Коэффициент спроса нагрузки": return ParameterType.ElectricalDemandFactor;
                case "Количество полюсов": return ParameterType.NumberOfPoles;
                case "Классификация нагрузок": return ParameterType.LoadClassification;
                case "Плотность (Трубопроводы)": return ParameterType.PipingDensity;
                case "Расход": return ParameterType.PipingFlow;
                case "Трение (Трубопроводы)": return ParameterType.PipingFriction;
                case "Давление (Трубопроводы)": return ParameterType.PipingPressure;
                case "Температура (Трубопроводы)": return ParameterType.PipingTemperature;
                case "Разность температур (Трубопроводы)": return ParameterType.PipingTemperatureDifference;
                case "Скорость (Трубопроводы)": return ParameterType.PipingVelocity;
                case "Динамическая вязкость (Трубопроводы)": return ParameterType.PipingViscosity;
                case "Размер трубы (PipeSize)": return ParameterType.PipeSize;
                case "Размер трубы (PipeDimension)": return ParameterType.PipeDimension;
                case "Шероховатость (Трубопроводы)": return ParameterType.PipingRoughness;
                case "Объем (Трубопроводы)": return ParameterType.PipingVolume;
                case "Уклон (Трубопроводы)": return ParameterType.PipingSlope;
                case "Толщина изоляции трубы": return ParameterType.PipeInsulationThickness;
                case "Масса (Трубопроводы)": return ParameterType.PipeMass;
                case "Масса на единицу длины (Трубопроводы)": return ParameterType.PipeMassPerUnitLength;
                case "Расход приборов": return ParameterType.FixtureUnit;
                case "Энергия (Энергия)": return ParameterType.HVACEnergy;
                case "Коэффициент теплопередачи": return ParameterType.HVACCoefficientOfHeatTransfer;
                case "Термостойкость": return ParameterType.HVACThermalResistance;
                case "Тепловая нагрузка": return ParameterType.HVACThermalMass;
                case "Теплопроводность": return ParameterType.HVACThermalConductivity;
                case "Удельная теплоемкость": return ParameterType.HVACSpecificHeat;
                case "Удельная теплоемкость парообразования": return ParameterType.HVACSpecificHeatOfVaporization;
                case "Проницаемость": return ParameterType.HVACPermeability;
                default: return ParameterType.Text;
            }
        }
#endif
#if Revit2023 || Debug2023
        public Category GetParameterTypeFromString(string dataType)
        {
            switch (dataType)
            { 
                default: return null;
            }
        }
#endif












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
                            { "NE", entry.Key },
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

        // Формирование словаря с параметрами из JSON-файла преднастроек
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

        //// Добавление новой панели параметров uniqueParameterField из JSON
        private void AddPanelParamFieldsJson(Dictionary<string, List<string>> allParamInInterfaceFromJsonDict)
        {
            StackPanel stackPanel = this.FindName("SP_panelParamFields") as StackPanel;

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
                    Width = 210,
                    Height = 25,
                    Padding = new Thickness(8, 4, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                foreach (var dataType in CategoryParameterDataTypes[allParamInInterfaceFromJsonList[3]])
                {
                    cbDataType.Items.Add(dataType);
                }

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
                        cbDataType.SelectedIndex = 3;
                    }
                    else
                    {
                        cbDataType.SelectedIndex = 0;
                    }
                };

                System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
                {
                    Width = 340,
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

                tbParamValue.GotFocus += TB_paramValue_GotFocus;

                tbParamValue.LostFocus += (s, ev) =>
                {
                    if (tbParamValue.Text == "")
                    {
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                    else
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117));
                    }
                };

                cbDataType.SelectionChanged += (s, ev) =>
                {
                    if (cbDataType.SelectedItem != null)
                    {
                        cbDataType.ToolTip = cbDataType.SelectedItem.ToString();
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра [{cbDataType.SelectedItem.ToString()}]";
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

                if (allParamInInterfaceFromJsonList[6] == "None")
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
            }
        }
    }
}