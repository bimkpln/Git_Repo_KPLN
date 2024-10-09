using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;


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

        public batchAddingParametersWindowFamily(UIApplication uiapp, string activeFamilyName, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            revitApp = uiapp.Application;
            _doc = uiapp.ActiveUIDocument?.Document;

            this.activeFamilyName = activeFamilyName;
            this.jsonFileSettingPath = jsonFileSettingPath;

            FillingComboBoxTypeInstance();
            FillingComboBoxGroupingName();
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
    }            
}

