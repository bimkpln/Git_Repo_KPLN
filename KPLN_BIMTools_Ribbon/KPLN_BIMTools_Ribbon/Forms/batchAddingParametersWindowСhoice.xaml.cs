using Autodesk.Revit.UI;
using System.Windows;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Linq;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Security.Cryptography;
using RevitServerAPILib;
using System.Data.Common;



namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowСhoice : Window
    {
        public batchAddingParametersWindowСhoice(UIApplication uiapp, string activeFamilyName)
        {        
            InitializeComponent();
            this.uiapp = uiapp;
            this.activeFamilyName = activeFamilyName;
            familyName.Text = activeFamilyName;
        }

        UIApplication uiapp;
        public string activeFamilyName;
        public string paramAction;
        public string paramType;
        public string jsonFileSettingPath;

        // Создание Dictionary с параметрами группирования для "Параметры группирования"
        static public Dictionary<string, BuiltInParameterGroup> CreateGroupingDictionary()
        {
            Dictionary<string, BuiltInParameterGroup> groupingDict = new Dictionary<string, BuiltInParameterGroup>
            {
                { "Аналитическая модель", BuiltInParameterGroup.PG_ANALYTICAL_MODEL },
                { "Видимость", BuiltInParameterGroup.PG_VISIBILITY },
                { "Второстепенный конец", BuiltInParameterGroup.PG_SECONDARY_END },
                { "Выравнивание аналитической модели", BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT },
                { "Геометрия разделения", BuiltInParameterGroup.PG_DIVISION_GEOMETRY },
                { "Графика", BuiltInParameterGroup.PG_GRAPHICS },
                { "Данные", BuiltInParameterGroup.PG_DATA },
                { "Зависимости", BuiltInParameterGroup.PG_CONSTRAINTS },
                { "Идентификация", BuiltInParameterGroup.PG_IDENTITY_DATA },
                { "Материалы и отделка", BuiltInParameterGroup.PG_MATERIALS },
                { "Механизмы", BuiltInParameterGroup.PG_MECHANICAL },
                { "Механизмы - Нагрузки", BuiltInParameterGroup.PG_MECHANICAL_LOADS },
                { "Механизмы - Расход", BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW },
                { "Моменты", BuiltInParameterGroup.PG_MOMENTS },
                { "Набор", BuiltInParameterGroup.PG_COUPLER_ARRAY },
                { "Набор арматурных стержней", BuiltInParameterGroup.PG_REBAR_ARRAY },
                { "Несущие конструкции", BuiltInParameterGroup.PG_STRUCTURAL },
                { "Общая легенда", BuiltInParameterGroup.PG_OVERALL_LEGEND },
                { "Общие", BuiltInParameterGroup.PG_GENERAL },
                { "Основной конец", BuiltInParameterGroup.PG_PRIMARY_END },
                { "Параметры IFC", BuiltInParameterGroup.PG_IFC },
                { "Прочее", BuiltInParameterGroup.INVALID },
                { "Размеры", BuiltInParameterGroup.PG_GEOMETRY },
                { "Расчет несущих конструкций", BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS },
                { "Расчет энергопотребления", BuiltInParameterGroup.PG_ENERGY_ANALYSIS },
                { "Редактирование формы перекрытия", BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT },
                { "Результат анализа", BuiltInParameterGroup.PG_ANALYSIS_RESULTS },
                { "Сантехника", BuiltInParameterGroup.PG_PLUMBING },
                { "Свойства модели", BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES },
                { "Свойства экологически чистого здания", BuiltInParameterGroup.PG_GREEN_BUILDING },
                { "Сегменты и соединительные детали", BuiltInParameterGroup.PG_SEGMENTS_FITTINGS },
                { "Силы", BuiltInParameterGroup.PG_FORCES },
                { "Система пожаротушения", BuiltInParameterGroup.PG_FIRE_PROTECTION },
                { "Слои", BuiltInParameterGroup.PG_REBAR_SYSTEM_LAYERS },
                { "Снятие связей/усилия для элемента", BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES },
                { "Стадии", BuiltInParameterGroup.PG_PHASING },
                { "Строительство", BuiltInParameterGroup.PG_CONSTRUCTION },
                { "Текст", BuiltInParameterGroup.PG_TEXT },
                { "Фотометрические", BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS },
                { "Шрифт заголовков", BuiltInParameterGroup.PG_TITLE },
#if Revit2023 || Debug2023
                { "Электросети (PG_ELECTRICAL_ENGINEERING)", BuiltInParameterGroup.PG_ELECTRICAL_ENGINEERING },
#endif
                { "Электросети (PG_ELECTRICAL)", BuiltInParameterGroup.PG_ELECTRICAL },
                { "Электросети - Нагрузки", BuiltInParameterGroup.PG_ELECTRICAL_LOADS },
                { "Электросети - Освещение", BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING },
                { "Электросети - Создание цепей", BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING }
            };

            return groupingDict;
        }

#if Revit2023 || Debug2023
        // Функция предопределения типа для ForgeTypeId 
        public string GetParamTypeName(ExternalDefinition def, ForgeTypeId value)
        {
            if (value == SpecTypeId.String.Text)
                return "Text";
            if (value == SpecTypeId.String.MultilineText)
                return "MultilineText";
            if (value == SpecTypeId.String.Url)
                return "URL";
            if (value == SpecTypeId.Int.Integer)
                return "Integer";
            if (value == SpecTypeId.Int.NumberOfPoles)
                return "NumberOfPoles";
            if (value == SpecTypeId.Boolean.YesNo)
                return "YesNo";
            if (value == SpecTypeId.Reference.Material)
                return "Material";
            if (value == SpecTypeId.Reference.LoadClassification)
                return "LoadClassification";
            if (value == SpecTypeId.Reference.Image)
                return "Image";

            return value.ToString();
        }
#endif


        // Функция соотношения параметра с типом данных и установки нужного кол-ва знаков после запятой.
        // Возвращает "yellow" - параметр пуст или указан неверно;
        // Возвращает "blue" - невозможно проверить значение параметра;
        // Возвращает "green" - значение параметра прошло проверку;
        // Возвращает "red" - значение параметра не прошло проверку;
        public string CheckingValueOfAParameter(System.Windows.Controls.ComboBox comboBox, System.Windows.Controls.TextBox textBox, string paramTypeName)
        {
            var textInField = textBox.Text;

            if (comboBox.SelectedItem == null)
            {
                return "yellow";
            }

            if (paramTypeName == "Image")
            {
                if (!string.IsNullOrEmpty(textInField))
                {
                    return "blue";
                }
            }

            if (paramTypeName == "Autodesk.Revit.DB.ForgeTypeId")
            {
                if (!string.IsNullOrEmpty(textInField))
                    if (double.TryParse(textInField, out double anyNumber))
                    {
                        return "green";
                    }
                    else
                    {
                        return "red";
                    }
            }

            if (paramTypeName == "Material")
            {
                List<string> materialNames = new FilteredElementCollector(uiapp.ActiveUIDocument.Document)
                                .OfClass(typeof(Material))
                                .Cast<Material>()
                                .Select(m => m.Name.ToLower())
                                .ToList();

                if (!string.IsNullOrEmpty(textInField) && materialNames.Contains(textInField.ToLower()))
                {
                    return "green";
                }
            }

            if (paramTypeName == "MultilineText" || paramTypeName == "Text" || paramTypeName =="URL")
            {
                if (!string.IsNullOrEmpty(textInField))
                {
                    return "green";
                }
            }

            if (paramTypeName == "YesNo")
            {
                if (textBox.Text == "0" || textBox.Text == "1")
                {
                    return "green";
                }
                else
                {
                    textBox.Text = "Необходимо указать: ``1`` - да; ``0`` - нет";
                    return "red";
                }
            }

            if (paramTypeName == "Integer")
            {
                if (textInField.Contains(","))
                {
                    return "red";
                }

                if (int.TryParse(textInField, out int resultInt))
                {
                    textBox.Text = resultInt.ToString();
                    return "green";
                }
            }

            if (paramTypeName == "NumberOfPoles")
            {
                if (int.TryParse(textInField, out int resultIntU) && resultIntU >= 1 && resultIntU <= 3)
                {
                    textBox.Text = resultIntU.ToString();
                    return "green";
                }
                else
                {
                    textBox.Text = "Необходимо указать: диапазон от 1 до 3";
                }
            }

            if (double.TryParse(textInField, out double resultDouble))
            {
                textBox.Text = resultDouble.ToString();
                return "green";
            }
            
            return "red";
        }

        // Функция соотношенияч типа данных со значением при добавлении параметра в семейство
        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, String parameterValue, String parameterValueDataType)
        {
            switch (parameterValueDataType)
            {
                case "Material":
                    Material material = new FilteredElementCollector(uiapp.ActiveUIDocument.Document)
                        .OfClass(typeof(Material))
                        .Cast<Material>()
                        .FirstOrDefault(m => m.Name.Equals(parameterValue));

                    if (material != null)
                    {
                        ElementId materialId = material.Id;

                        familyManager.Set(familyParam, materialId);
                    }
                    break;

                case "Text":
                case "MultilineText":
                case "URL":
                    familyManager.Set(familyParam, parameterValue);
                    break;

                case "Integer":
                case "YesNo":
                case "NumberOfPoles":
                    if (int.TryParse(parameterValue, out int intBoolValue))
                    {
                        familyManager.Set(familyParam, intBoolValue);
                    }
                    break;

#if Revit2020 || Debug2020
                case "Image":
                    string imagePath = parameterValue;

                    FilteredElementCollector collector = new FilteredElementCollector(uiapp.ActiveUIDocument.Document)
                        .OfClass(typeof(ImageType));

                    ImageType imageType = collector
                        .Cast<ImageType>()
                        .FirstOrDefault(img => img.Name.Equals(Path.GetFileName(imagePath), StringComparison.OrdinalIgnoreCase));

                    if (imageType != null)
                    {
                        familyManager.Set(familyParam, imageType.Id);
                    }
                    else
                    {
                        ImageType newImageTypeOld = ImageType.Create(uiapp.ActiveUIDocument.Document, imagePath);
                        familyManager.Set(familyParam, newImageTypeOld.Id);
                    }
                    break;

                default:
                    if (double.TryParse(parameterValue, out double millimetersValue))
                    {
                        UnitType unitType = familyParam.Definition.UnitType;
                        DisplayUnitType displayUnitType = uiapp.ActiveUIDocument.Document.GetUnits().GetFormatOptions(unitType).DisplayUnits;

                        double convertedValue = UnitUtils.ConvertToInternalUnits(millimetersValue, displayUnitType);
                        familyManager.Set(familyParam, convertedValue);
                    }
                    break;
#endif
#if Revit2023 || Debug2023
                case "Image":
                    string imagePath = parameterValue;

                    FilteredElementCollector collector = new FilteredElementCollector(uiapp.ActiveUIDocument.Document)
                        .OfClass(typeof(ImageType));

                    ImageType imageType = collector
                        .Cast<ImageType>()
                        .FirstOrDefault(img => img.Name.Equals(Path.GetFileName(imagePath), StringComparison.OrdinalIgnoreCase));

                    if (imageType != null)
                    {
                        familyManager.Set(familyParam, imageType.Id);
                    }
                    else
                    {
                        ImageTypeOptions options = new ImageTypeOptions(imagePath, false, ImageTypeSource.Import);
                        ImageType newImageType = ImageType.Create(uiapp.ActiveUIDocument.Document, options);
                        familyManager.Set(familyParam, newImageType.Id);
                    }
                    break;

                default:
                    if (double.TryParse(parameterValue, out double millimetersValue))
                    {
                        ForgeTypeId forgeTypeId = familyParam.Definition.GetDataType();
                        FormatOptions formatOptions = uiapp.ActiveUIDocument.Document.GetUnits().GetFormatOptions(forgeTypeId);

                        double convertedValue = UnitUtils.ConvertToInternalUnits(millimetersValue, formatOptions.GetUnitTypeId());
                        familyManager.Set(familyParam, convertedValue);
                    }
                    break;
#endif
            }
        }
#endif
















        //// XAML. Пакетное добавление общих параметров
        private void Button_NewGeneralParam(object sender, RoutedEventArgs e)
        {
            jsonFileSettingPath = "";

            var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
        }

        //// XAML. Пакетное добавление кастомных параметров семейства
        private void Button_NewFamilyParam(object sender, RoutedEventArgs e)
        {
            var window = new batchAddingParametersWindowFamily(uiapp, activeFamilyName, jsonFileSettingPath);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();
        }

        //// XAML. Загрузка XAML-настройки
        private void Button_LoadParam(object sender, RoutedEventArgs e)
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
                    var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
                    var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                    window.ShowDialog();
                }
                else if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["quantity"] != null && item["parameterName"] != null && item["instance"] != null && item["categoryType"] != null 
                        && item["dataType"] != null && item["grouping"] != null && item["parameterValue"] != null && item["comment"] != null))
                {
                    var window = new batchAddingParametersWindowFamily(uiapp, activeFamilyName, jsonFileSettingPath);
                    var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                    window.ShowDialog();
                }
                else{
                    System.Windows.Forms.MessageBox.Show("Ваш JSON-файл не является файлом преднастроек или повреждён. Пожалуйста, выберите другой файл.", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }
    }
}
