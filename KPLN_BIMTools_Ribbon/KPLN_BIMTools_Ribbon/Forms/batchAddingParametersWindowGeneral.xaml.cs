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
                newParamList[nameParameter].Add(entry["parameterValue"]);
                newParamList[nameParameter].Add(entry["parameterValueDataType"]);
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

                            if (textBoxValue == "" || textBoxValue.Contains("При необходимости, вы можете указать значение параметра"))
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

                        interfaceParamDict.Add(stackPanelKey, contentsOfDataFields);
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

        // Функция соотношения параметра с типом данных.
        // Возвращает "red" - значение параметра не прошло проверку;
        // Возвращает "blue" - невозможно проверить значение параметра;
        // Возвращает "yellow" - параметр пуст или указан неверно;
        // Возвращает "green" - значение параметра прошло проверку;

        public string CheckingValueOfAParameter(System.Windows.Controls.ComboBox ComboBox, System.Windows.Controls.TextBox textBox, ParameterType paramType)
        {
            var textInField = textBox.Text;

            if (ComboBox.SelectedItem == null)
            {
                return "yellow";
            }

            if (paramType == ParameterType.MultilineText || paramType == ParameterType.Text || paramType == ParameterType.URL)
            {
                if (!string.IsNullOrEmpty(textInField))
                {
                    return "green";
                }
            }

            if (paramType == ParameterType.YesNo)
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

            if (paramType == ParameterType.Integer)
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

            if (paramType == ParameterType.ElectricalCableTraySize || paramType == ParameterType.ElectricalConduitSize || paramType == ParameterType.ForceLengthPerAngle || paramType == ParameterType.HVACDuctSize || paramType == ParameterType.HVACDuctLiningThickness 
                || paramType == ParameterType.HVACDuctInsulationThickness || paramType == ParameterType.HVACEnergy || paramType == ParameterType.LinearForceLengthPerAngle || paramType == ParameterType.PipeSize || paramType == ParameterType.PipeInsulationThickness 
                || paramType == ParameterType.PipingVolume)
            {
                if (double.TryParse(textInField, out double resultDouble_1))
                {
                    textBox.Text = resultDouble_1.ToString("F1");
                    return "green";
                }
            }

            if (paramType == ParameterType.Angle || paramType == ParameterType.AreaForce || paramType == ParameterType.AreaForcePerLength || paramType == ParameterType.BarDiameter || paramType == ParameterType.ColorTemperature || paramType == ParameterType.CrackWidth 
                || paramType == ParameterType.Currency || paramType == ParameterType.DisplacementDeflection || paramType == ParameterType.ElectricalApparentPower || paramType == ParameterType.ElectricalCurrent || paramType == ParameterType.ElectricalEfficacy 
                || paramType == ParameterType.ElectricalIlluminance || paramType == ParameterType.ElectricalLuminance || paramType == ParameterType.ElectricalFrequency || paramType == ParameterType.ElectricalLuminousIntensity || paramType == ParameterType.ElectricalLuminousFlux 
                || paramType == ParameterType.ElectricalPotential || paramType == ParameterType.ElectricalPower || paramType == ParameterType.ElectricalPowerDensity || paramType == ParameterType.ElectricalTemperature || paramType == ParameterType.ElectricalTemperatureDifference 
                || paramType == ParameterType.ElectricalWattage || paramType == ParameterType.Force || paramType == ParameterType.ForcePerLength || paramType == ParameterType.HVACAirflow || paramType == ParameterType.HVACAirflowDensity 
                || paramType == ParameterType.HVACAirflowDividedByCoolingLoad || paramType == ParameterType.HVACAirflowDividedByVolume || paramType == ParameterType.HVACAreaDividedByCoolingLoad || paramType == ParameterType.HVACCoolingLoad 
                || paramType == ParameterType.HVACCoolingLoadDividedByArea || paramType == ParameterType.HVACCoolingLoadDividedByVolume || paramType == ParameterType.HVACCrossSection || paramType == ParameterType.HVACHeatGain || paramType == ParameterType.HVACHeatingLoad 
                || paramType == ParameterType.HVACHeatingLoadDividedByArea || paramType == ParameterType.HVACHeatingLoadDividedByVolume || paramType == ParameterType.HVACPower || paramType == ParameterType.HVACPowerDensity || paramType == ParameterType.HVACPressure 
                || paramType == ParameterType.HVACRoughness || paramType == ParameterType.HVACTemperature || paramType == ParameterType.HVACTemperatureDifference || paramType == ParameterType.HVACThermalMass || paramType == ParameterType.HVACVelocity 
                || paramType == ParameterType.Length || paramType == ParameterType.LinearForce || paramType == ParameterType.LinearForcePerLength || paramType == ParameterType.LinearMoment || paramType == ParameterType.MassPerUnitArea || paramType == ParameterType.Moment 
                || paramType == ParameterType.MomentOfInertia || paramType == ParameterType.PipeDimension || paramType == ParameterType.PipingFlow || paramType == ParameterType.PipingPressure || paramType == ParameterType.PipingTemperature 
                || paramType == ParameterType.PipingTemperatureDifference || paramType == ParameterType.PipingVelocity || paramType == ParameterType.ReinforcementArea || paramType == ParameterType.ReinforcementCover ||  paramType == ParameterType.ReinforcementLength 
                || paramType == ParameterType.ReinforcementSpacing || paramType == ParameterType.ReinforcementVolume || paramType == ParameterType.SectionProperty || paramType == ParameterType.SectionArea || paramType == ParameterType.SectionDimension 
                || paramType == ParameterType.SectionModulus || paramType == ParameterType.StructuralFrequency || paramType == ParameterType.StructuralVelocity || paramType == ParameterType.UnitWeight || paramType == ParameterType.Weight 
                || paramType == ParameterType.WeightPerUnitLength || paramType == ParameterType.WireSize)
            {
                if (double.TryParse(textInField, out double resultDouble_2))
                {
                    textBox.Text = resultDouble_2.ToString("F2");
                    return "green";
                }
            }

            if (paramType == ParameterType.Slope)
            {
                if (double.TryParse(textInField, out double resultDouble_2) && resultDouble_2 >= -89.98 && resultDouble_2 <= 89.98)
                {
                    textBox.Text = resultDouble_2.ToString("F2");
                    return "green";
                }
                else
                {
                    textBox.Text = "Необходимо указать: диапазон от -89.98 до 89.98";
                }
            }

            if (paramType == ParameterType.Area || paramType == ParameterType.Acceleration || paramType == ParameterType.Energy || paramType == ParameterType.Mass || paramType == ParameterType.MassPerUnitLength || paramType == ParameterType.Period 
                || paramType == ParameterType.PipeMass || paramType == ParameterType.PipeMassPerUnitLength || paramType == ParameterType.PipingRoughness || paramType == ParameterType.Pulsation || paramType == ParameterType.ReinforcementAreaPerUnitLength
                || paramType == ParameterType.Rotation || paramType == ParameterType.Speed || paramType == ParameterType.SurfaceArea || paramType == ParameterType.TimeInterval || paramType == ParameterType.Volume || paramType == ParameterType.WarpingConstant)
            {
                if (double.TryParse(textInField, out double resultDouble_3))
                {
                    textBox.Text = resultDouble_3.ToString("F3");
                    return "green";
                }
            }

            if (paramType == ParameterType.ElectricalDemandFactor || paramType == ParameterType.ElectricalResistivity
                || paramType == ParameterType.HVACCoefficientOfHeatTransfer || paramType == ParameterType.HVACAreaDividedByHeatingLoad || paramType == ParameterType.HVACFactor || paramType == ParameterType.HVACFriction || paramType == ParameterType.HVACPermeability
                || paramType == ParameterType.HVACThermalConductivity || paramType == ParameterType.HVACSlope || paramType == ParameterType.HVACSpecificHeat || paramType == ParameterType.HVACSpecificHeatOfVaporization || paramType == ParameterType.HVACViscosity 
                || paramType == ParameterType.PipingFriction || paramType == ParameterType.HVACThermalResistance || paramType == ParameterType.PipingSlope || paramType == ParameterType.PipingViscosity)
            {
                if (double.TryParse(textInField, out double resultDouble_4))
                {
                    textBox.Text = resultDouble_4.ToString("F4");
                    return "green";
                }
            }

            if (paramType == ParameterType.FixtureUnit || paramType == ParameterType.HVACDensity || paramType == ParameterType.MassDensity || paramType == ParameterType.Number || paramType == ParameterType.PipingDensity || paramType == ParameterType.Stress)
            {
                if (double.TryParse(textInField, out double resultDouble_6))
                {
                    textBox.Text = resultDouble_6.ToString("F6");
                    return "green";
                }
            }

            if (paramType == ParameterType.ThermalExpansion)
            {
                if (double.TryParse(textInField, out double resultDouble_8))
                {
                    textBox.Text = resultDouble_8.ToString("F8");
                    return "green";
                }
            }

            /////////////////////////////////////////////////////////////////////////////////////////// Тест
            if (paramType == ParameterType.Image || paramType == ParameterType.FamilyType || paramType == ParameterType.Material || paramType == ParameterType.LoadClassification)
            {
                if (!string.IsNullOrEmpty(textInField))
                {
                    return "blue";
                }
            }

            return "red";
        }





        // Функция соотношенияч типа данных со значением при добавлении параметра в семейство
        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, String parameterValue, String parameterValueDataType)
        {
            switch (parameterValueDataType)
            {
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                // Это пока добавляется, как текст (потом нужно что-то думать)
                case "Image":
                case "FamilyType":
                case "Material":
                case "LoadClassification":
                    familyManager.Set(familyParam, parameterValue);
                    break;
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                //////////////////
                /////////////////////////////////////////////////////////////////////////////////////////// Тест

                case "HVACPermeability":
                    if (double.TryParse(parameterValue, out double nanogramsPerPascalSecondSquareMeterValue))
                    {
                        double convertedNanogramsPerPascalSecondSquareMeter = UnitUtils.ConvertToInternalUnits(nanogramsPerPascalSecondSquareMeterValue, DisplayUnitType.DUT_NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER);
                        familyManager.Set(familyParam, convertedNanogramsPerPascalSecondSquareMeter);
                    }
                    break;


                /// 
                /// Типы без конверсии
                /// 
                case "Text":
                case "MultilineText":
                case "URL":
                    familyManager.Set(familyParam, parameterValue);
                    break;

                case "Integer":
                case "NumberOfPoles":
                case "YesNo":
                    if (int.TryParse(parameterValue, out int intBoolValue))
                    {
                        familyManager.Set(familyParam, intBoolValue);
                    }
                    break;

                case "ColorTemperature":
                case "Currency":
                case "ElectricalCurrent":
                case "ElectricalFrequency":
                case "ElectricalLuminousFlux":
                case "ElectricalLuminousIntensity":
                case "ElectricalPowerDensity":
                case "ElectricalTemperatureDifference":
                case "FixtureUnit":
                case "HVACCoolingLoadDividedByArea":
                case "HVACHeatingLoadDividedByArea":              
                case "HVACCoefficientOfHeatTransfer":
                case "HVACThermalResistance":
                case "HVACTemperatureDifference":              
                case "HVACPowerDensity":                
                case "Mass":
                case "Number":
                case "Period":
                case "PipeMass":
                case "PipingTemperatureDifference":                
                case "Pulsation":
                case "Rotation":
                case "StructuralFrequency":
                case "TimeInterval":
                case "ThermalExpansion":
                    if (double.TryParse(parameterValue, out double resultNumber))
                    {
                        familyManager.Set(familyParam, resultNumber);
                    }
                    break;

                /// 
                /// Типы с прямой конверсией
                /// 
                case "BarDiameter":
                case "CrackWidth":
                case "ElectricalCableTraySize":
                case "ElectricalConduitSize":               
                case "HVACDuctSize":
                case "HVACDuctLiningThickness":
                case "HVACDuctInsulationThickness":
                case "HVACRoughness":
                case "Length":
                case "PipeDimension":               
                case "PipeInsulationThickness":
                case "PipingRoughness":
                case "PipeSize":
                case "ReinforcementCover":
                case "ReinforcementLength":
                case "ReinforcementSpacing":
                case "WireSize":
                    if (double.TryParse(parameterValue, out double millimetersValue))
                    {
                        double convertedMillimeters = UnitUtils.ConvertToInternalUnits(millimetersValue, DisplayUnitType.DUT_MILLIMETERS);
                        familyManager.Set(familyParam, convertedMillimeters);
                    }
                    break;

                case "HVACCrossSection":
                    if (double.TryParse(parameterValue, out double squareMillimetersValue))
                    {
                        double convertedSquareMillimeters = UnitUtils.ConvertToInternalUnits(squareMillimetersValue, DisplayUnitType.DUT_SQUARE_MILLIMETERS);
                        familyManager.Set(familyParam, convertedSquareMillimeters);
                    }
                    break;

                case "DisplacementDeflection":
                case "SectionDimension":                
                case "SectionProperty":
                    if (double.TryParse(parameterValue, out double centimetrsValue))
                    {
                        double convertedCentimetrs = UnitUtils.ConvertToInternalUnits(centimetrsValue, DisplayUnitType.DUT_CENTIMETERS);
                        familyManager.Set(familyParam, convertedCentimetrs);
                    }
                    break;

                case "ReinforcementArea":
                case "SectionArea":
                    if (double.TryParse(parameterValue, out double squareCentimetersValue))
                    {
                        double convertedSquareCentimeters = UnitUtils.ConvertToInternalUnits(squareCentimetersValue, DisplayUnitType.DUT_SQUARE_CENTIMETERS);
                        familyManager.Set(familyParam, convertedSquareCentimeters);
                    }
                    break;

                case "ReinforcementVolume":
                case "SectionModulus":
                    if (double.TryParse(parameterValue, out double cubicCentimetersValue))
                    {
                        double convertedCubicCentimeters = UnitUtils.ConvertToInternalUnits(cubicCentimetersValue, DisplayUnitType.DUT_CUBIC_CENTIMETERS);
                        familyManager.Set(familyParam, convertedCubicCentimeters);
                    }
                    break;

                case "MomentOfInertia":
                    if (double.TryParse(parameterValue, out double centimetrsP4Value))
                    {
                        double convertedCentimetrsP6 = UnitUtils.ConvertToInternalUnits(centimetrsP4Value, DisplayUnitType.DUT_CENTIMETERS_TO_THE_FOURTH_POWER);
                        familyManager.Set(familyParam, convertedCentimetrsP6);
                    }
                    break;

                case "WarpingConstant":
                    if (double.TryParse(parameterValue, out double centimetrsP6Value))
                    {
                        double convertedCentimetrsP6 = UnitUtils.ConvertToInternalUnits(centimetrsP6Value, DisplayUnitType.DUT_CENTIMETERS_TO_THE_SIXTH_POWER);
                        familyManager.Set(familyParam, convertedCentimetrsP6);
                    }
                    break;

                case "ReinforcementAreaPerUnitLength":
                    if (double.TryParse(parameterValue, out double squareCentimetersPerMeterValue))
                    {
                        double convertedSquareCentimetersPerMeter = UnitUtils.ConvertToInternalUnits(squareCentimetersPerMeterValue, DisplayUnitType.DUT_SQUARE_CENTIMETERS_PER_METER);
                        familyManager.Set(familyParam, convertedSquareCentimetersPerMeter);
                    }
                    break;

                case "Area":
                    if (double.TryParse(parameterValue, out double sqMetersValue))
                    {
                        double convertedSqMeters = UnitUtils.ConvertToInternalUnits(sqMetersValue, DisplayUnitType.DUT_SQUARE_METERS);
                        familyManager.Set(familyParam, convertedSqMeters);
                    }
                    break;

                case "Volume":
                    if (double.TryParse(parameterValue, out double cubMetersValue))
                    {
                        double convertedCubMeters = UnitUtils.ConvertToInternalUnits(cubMetersValue, DisplayUnitType.DUT_CUBIC_METERS);
                        familyManager.Set(familyParam, convertedCubMeters);

                    }
                    break;

                case "HVACVelocity":
                case "PipingVelocity":
                case "StructuralVelocity":
                    if (double.TryParse(parameterValue, out double meterPerSecValue))
                    {

                        double convertedMeterPerSec = UnitUtils.ConvertToInternalUnits(meterPerSecValue, DisplayUnitType.DUT_METERS_PER_SECOND);
                        familyManager.Set(familyParam, convertedMeterPerSec);
                    }
                    break;

                case "Acceleration":
                    if (double.TryParse(parameterValue, out double meterPerSecSquaredValue))
                    {
                        double convertedMeterPerSecSquared = UnitUtils.ConvertToInternalUnits(meterPerSecSquaredValue, DisplayUnitType.DUT_METERS_PER_SECOND_SQUARED);
                        familyManager.Set(familyParam, convertedMeterPerSecSquared);
                    }
                    break;

                case "SurfaceArea":
                    if (double.TryParse(parameterValue, out double sqareMetersPerMeterValue))
                    {
                        double convertedSqareMetersPerMeter = UnitUtils.ConvertToInternalUnits(sqareMetersPerMeterValue, DisplayUnitType.DUT_SQUARE_METERS_PER_METER);
                        familyManager.Set(familyParam, convertedSqareMetersPerMeter);
                    }
                    break;
                
                case "HVACAreaDividedByCoolingLoad":
                case "HVACAreaDividedByHeatingLoad":
                    if (double.TryParse(parameterValue, out double squareMeterPerKilowattsValue))
                    {
                        double convertedSquareMeterPerKilowatts = UnitUtils.ConvertToInternalUnits(squareMeterPerKilowattsValue, DisplayUnitType.DUT_SQUARE_METERS_PER_KILOWATTS);
                        familyManager.Set(familyParam, convertedSquareMeterPerKilowatts);
                    }
                    break;

                case "HVACAirflowDividedByCoolingLoad":
                    if (double.TryParse(parameterValue, out double cubicMeterPerSecondValue))
                    {
                        double convertedCubicMeterPerSecond = UnitUtils.ConvertToInternalUnits(cubicMeterPerSecondValue, DisplayUnitType.DUT_CUBIC_METERS_PER_SECOND);
                        familyManager.Set(familyParam, convertedCubicMeterPerSecond);
                    }
                    break;

                case "Speed":
                    if (double.TryParse(parameterValue, out double kilMeterPerhHourValue))
                    {
                        double convertedKilMeterPerhHour = UnitUtils.ConvertToInternalUnits(kilMeterPerhHourValue, DisplayUnitType.DUT_KILOMETERS_PER_HOUR);
                        familyManager.Set(familyParam, convertedKilMeterPerhHour);
                    }
                    break;

                case "ElectricalDemandFactor":
                case "HVACFactor":
                case "HVACSlope":
                case "PipingSlope":
                    if (double.TryParse(parameterValue, out double percentageValue))
                    {
                        double convertedPercentage = UnitUtils.ConvertToInternalUnits(percentageValue, DisplayUnitType.DUT_PERCENTAGE);
                        familyManager.Set(familyParam, convertedPercentage);
                    }
                    break;

                case "Angle":
                    if (double.TryParse(parameterValue, out double decDegreesValue))
                    {
                        double convertedDecDegrees = UnitUtils.ConvertToInternalUnits(decDegreesValue, DisplayUnitType.DUT_DECIMAL_DEGREES);
                        familyManager.Set(familyParam, convertedDecDegrees);
                    }
                    break;

                case "Slope":
                    if (double.TryParse(parameterValue, out double slopeDegreesValue))
                    {
                        double convertedSlopeDegrees = UnitUtils.ConvertToInternalUnits(slopeDegreesValue, DisplayUnitType.DUT_SLOPE_DEGREES);
                        familyManager.Set(familyParam, convertedSlopeDegrees);
                    }
                    break;

                case "HVACEnergy":
                    if (double.TryParse(parameterValue, out double joulesValue))
                    {
                        double convertedJoules= UnitUtils.ConvertToInternalUnits(joulesValue, DisplayUnitType.DUT_JOULES);
                        familyManager.Set(familyParam, convertedJoules);
                    }
                    break;

                case "HVACSpecificHeatOfVaporization":
                    if (double.TryParse(parameterValue, out double joulesPerGramValue))
                    {
                        double convertedJoulesPerGram = UnitUtils.ConvertToInternalUnits(joulesPerGramValue, DisplayUnitType.DUT_JOULES_PER_GRAM);
                        familyManager.Set(familyParam, convertedJoulesPerGram);
                    }
                    break;

                case "HVACSpecificHeat":
                    if (double.TryParse(parameterValue, out double joulesPerKilogramCelsiusValue))
                    {
                        double convertedJoulesPerKilogramCelsius = UnitUtils.ConvertToInternalUnits(joulesPerKilogramCelsiusValue, DisplayUnitType.DUT_JOULES_PER_KILOGRAM_CELSIUS);
                        familyManager.Set(familyParam, convertedJoulesPerKilogramCelsius);
                    }
                    break;

                case "Energy":
                    if (double.TryParse(parameterValue, out double kilojoulesValue))
                    {
                        double convertedKilojoules = UnitUtils.ConvertToInternalUnits(kilojoulesValue, DisplayUnitType.DUT_KILOJOULES);
                        familyManager.Set(familyParam, convertedKilojoules);
                    }
                    break;

                case "HVACThermalMass":
                    if (double.TryParse(parameterValue, out double kilojoulesPerKelvinValue))
                    {
                        double convertedKilojoulesPerKelvin = UnitUtils.ConvertToInternalUnits(kilojoulesPerKelvinValue, DisplayUnitType.DUT_KILOJOULES_PER_KELVIN);
                        familyManager.Set(familyParam, convertedKilojoulesPerKelvin);
                    }
                    break;

                case "ElectricalPotential":
                    if (double.TryParse(parameterValue, out double voltsValue))
                    {
                        double convertedVolts = UnitUtils.ConvertToInternalUnits(voltsValue, DisplayUnitType.DUT_VOLTS);
                        familyManager.Set(familyParam, convertedVolts);
                    }
                    break;

                case "ElectricalApparentPower":
                    if (double.TryParse(parameterValue, out double voltAmperesValue))
                    {
                        double convertedVoltAmperes = UnitUtils.ConvertToInternalUnits(voltAmperesValue, DisplayUnitType.DUT_VOLT_AMPERES);
                        familyManager.Set(familyParam, convertedVoltAmperes);
                    }
                    break;

                case "ElectricalPower":
                case "HVACCoolingLoad":
                case "HVACHeatGain":
                case "HVACHeatingLoad":
                case "HVACPower":
                case "ElectricalWattage":                
                    if (double.TryParse(parameterValue, out double wattsValue))
                    {
                        double convertedWatts = UnitUtils.ConvertToInternalUnits(wattsValue, DisplayUnitType.DUT_WATTS);
                        familyManager.Set(familyParam, convertedWatts);
                    }
                    break;

                case "HVACHeatingLoadDividedByVolume":
                case "HVACCoolingLoadDividedByVolume":
                    if (double.TryParse(parameterValue, out double watsPerCubicMeterValue))
                    {
                        double convertedWatsPerCubicMeter = UnitUtils.ConvertToInternalUnits(watsPerCubicMeterValue, DisplayUnitType.DUT_WATTS_PER_CUBIC_METER);
                        familyManager.Set(familyParam, convertedWatsPerCubicMeter);
                    }
                    break;

                case "HVACThermalConductivity":
                    if (double.TryParse(parameterValue, out double wattsPerMeterKelvinValue))
                    {
                        double convertedWattsPerMeterKelvin = UnitUtils.ConvertToInternalUnits(wattsPerMeterKelvinValue, DisplayUnitType.DUT_WATTS_PER_METER_KELVIN);
                        familyManager.Set(familyParam, convertedWattsPerMeterKelvin);
                    }
                    break;

                case "ElectricalResistivity":
                    if (double.TryParse(parameterValue, out double ohmMetersValue))
                    {
                        double convertedOhmMeters = UnitUtils.ConvertToInternalUnits(ohmMetersValue, DisplayUnitType.DUT_OHM_METERS);
                        familyManager.Set(familyParam, convertedOhmMeters);
                    }
                    break;

                case "HVACPressure":
                case "PipingPressure":
                    if (double.TryParse(parameterValue, out double pascValue))
                    {
                        double convertedPasc = UnitUtils.ConvertToInternalUnits(pascValue, DisplayUnitType.DUT_PASCALS);
                        familyManager.Set(familyParam, convertedPasc);

                    }
                    break;

                case "HVACFriction":
                case "PipingFriction":
                    if (double.TryParse(parameterValue, out double pascalPerMeterValue))
                    {
                        double convertedPascalPerMeter = UnitUtils.ConvertToInternalUnits(pascalPerMeterValue, DisplayUnitType.DUT_PASCALS_PER_METER);
                        familyManager.Set(familyParam, convertedPascalPerMeter);
                    }
                    break;

                case "HVACViscosity":
                case "PipingViscosity":
                    if (double.TryParse(parameterValue, out double pascalSecondsValue))
                    {
                        double convertedPascalSeconds = UnitUtils.ConvertToInternalUnits(pascalSecondsValue, DisplayUnitType.DUT_PASCAL_SECONDS);
                        familyManager.Set(familyParam, convertedPascalSeconds);
                    }
                    break;

                case "Stress":
                    if (double.TryParse(parameterValue, out double megaPascalValue))
                    {
                        double convertedMegaPascal = UnitUtils.ConvertToInternalUnits(megaPascalValue, DisplayUnitType.DUT_MEGAPASCALS);
                        familyManager.Set(familyParam, convertedMegaPascal);
                    }
                    break;

                case "Force":
                case "Weight":
                    if (double.TryParse(parameterValue, out double kilonewtonsValue))
                    {
                        double convertedKilonewtons = UnitUtils.ConvertToInternalUnits(kilonewtonsValue, DisplayUnitType.DUT_KILONEWTONS);
                        familyManager.Set(familyParam, convertedKilonewtons);
                    }
                    break;

                case "Moment":
                    if (double.TryParse(parameterValue, out double kilonewtonMeterValue))
                    {
                        double convertedKilonewtonMeter = UnitUtils.ConvertToInternalUnits(kilonewtonMeterValue, DisplayUnitType.DUT_KILONEWTON_METERS);
                        familyManager.Set(familyParam, convertedKilonewtonMeter);
                    }
                    break;

                case "ForcePerLength":
                case "LinearForce":
                    if (double.TryParse(parameterValue, out double kilonewtonPerMeterValue))
                    {
                        double convertedKilonewtonPerMeter = UnitUtils.ConvertToInternalUnits(kilonewtonPerMeterValue, DisplayUnitType.DUT_KILONEWTONS_PER_METER);
                        familyManager.Set(familyParam, convertedKilonewtonPerMeter);
                    }
                    break;

                case "LinearMoment":
                    if (double.TryParse(parameterValue, out double kilonewtonMeterPerMeterValue))
                    {
                        double convertedKilonewtonMeterPerMeter = UnitUtils.ConvertToInternalUnits(kilonewtonMeterPerMeterValue, DisplayUnitType.DUT_KILONEWTON_METERS_PER_METER);
                        familyManager.Set(familyParam, convertedKilonewtonMeterPerMeter);
                    }
                    break;

                case "AreaForce":
                case "LinearForcePerLength":
                    if (double.TryParse(parameterValue, out double kilonewtonsPerSquareMeterValue))
                    {
                        double convertedKilonewtonsPerSquareMeter = UnitUtils.ConvertToInternalUnits(kilonewtonsPerSquareMeterValue, DisplayUnitType.DUT_KILONEWTONS_PER_SQUARE_METER);
                        familyManager.Set(familyParam, convertedKilonewtonsPerSquareMeter);
                    }
                    break;

                case "AreaForcePerLength":
                case "UnitWeight":
                    if (double.TryParse(parameterValue, out double kilonewtonsPerCubicMeterValue))
                    {
                        double convertedKilonewtonsPerCubicMeter = UnitUtils.ConvertToInternalUnits(kilonewtonsPerCubicMeterValue, DisplayUnitType.DUT_KILONEWTONS_PER_CUBIC_METER);
                        familyManager.Set(familyParam, convertedKilonewtonsPerCubicMeter);
                    }
                    break;

                case "ForceLengthPerAngle":
                    if (double.TryParse(parameterValue, out double kilonewtonMetersPerDegreeValue))
                    {
                        double convertedKilonewtonMetersPerDegree = UnitUtils.ConvertToInternalUnits(kilonewtonMetersPerDegreeValue, DisplayUnitType.DUT_KILONEWTON_METERS_PER_DEGREE);
                        familyManager.Set(familyParam, convertedKilonewtonMetersPerDegree);
                    }
                    break;

                case "LinearForceLengthPerAngle":
                    if (double.TryParse(parameterValue, out double kilonewtonMetersPerDegreePerMeterValue))
                    {
                        double convertedKilonewtonMetersPerDegreePerMeter = UnitUtils.ConvertToInternalUnits(kilonewtonMetersPerDegreePerMeterValue, DisplayUnitType.DUT_KILONEWTON_METERS_PER_DEGREE_PER_METER);
                        familyManager.Set(familyParam, convertedKilonewtonMetersPerDegreePerMeter);
                    }
                    break;

                case "HVACDensity":
                case "MassDensity":
                case "PipingDensity":
                    if (double.TryParse(parameterValue, out double kgPerCubMeterValue))
                    {
                        double convertedKgPerCubMeter = UnitUtils.ConvertToInternalUnits(kgPerCubMeterValue, DisplayUnitType.DUT_KILOGRAMS_PER_CUBIC_METER);
                        familyManager.Set(familyParam, convertedKgPerCubMeter);
                    }
                    break;

                case "WeightPerUnitLength":
                    if (double.TryParse(parameterValue, out double kilForcePerMeterValue))
                    {
                        double convertedKilForcePerMeter = UnitUtils.ConvertToInternalUnits(kilForcePerMeterValue, DisplayUnitType.DUT_KILOGRAMS_FORCE_PER_METER);
                        familyManager.Set(familyParam, convertedKilForcePerMeter);
                    }
                    break;

                case "MassPerUnitLength":
                case "PipeMassPerUnitLength":
                    if (double.TryParse(parameterValue, out double killogramMassPerMeterValue))
                    {
                        double convertedKillogramMassPerMeter = UnitUtils.ConvertToInternalUnits(killogramMassPerMeterValue, DisplayUnitType.DUT_KILOGRAMS_MASS_PER_METER);
                        familyManager.Set(familyParam, convertedKillogramMassPerMeter);
                    }
                    break;

                case "MassPerUnitArea":
                    if (double.TryParse(parameterValue, out double killogramMassPerSquareMeterValue))
                    {
                        double convertedKillogramMassPerSquareMeter = UnitUtils.ConvertToInternalUnits(killogramMassPerSquareMeterValue, DisplayUnitType.DUT_KILOGRAMS_MASS_PER_SQUARE_METER);
                        familyManager.Set(familyParam, convertedKillogramMassPerSquareMeter);
                    }
                    break;

                case "PipingVolume":
                    if (double.TryParse(parameterValue, out double litersValue))
                    {
                        double convertedLiters = UnitUtils.ConvertToInternalUnits(litersValue, DisplayUnitType.DUT_LITERS);
                        familyManager.Set(familyParam, convertedLiters);
                    }
                    break;

                case "HVACAirflow":
                case "PipingFlow":
                    if (double.TryParse(parameterValue, out double litersPerSecondValue))
                    {
                        double convertedLitersPerSecond = UnitUtils.ConvertToInternalUnits(litersPerSecondValue, DisplayUnitType.DUT_LITERS_PER_SECOND);
                        familyManager.Set(familyParam, convertedLitersPerSecond);
                    }
                    break;

                case "HVACAirflowDensity":
                    if (double.TryParse(parameterValue, out double litersPerSecondSquareMeterValue))
                    {
                        double convertedLitersPerSecondSquareMeter = UnitUtils.ConvertToInternalUnits(litersPerSecondSquareMeterValue, DisplayUnitType.DUT_LITERS_PER_SECOND_SQUARE_METER);
                        familyManager.Set(familyParam, convertedLitersPerSecondSquareMeter);
                    }
                    break;

                case "HVACAirflowDividedByVolume":
                    if (double.TryParse(parameterValue, out double litersPerSecondCubicMeterValue))
                    {
                        double convertedLitersPerSecondCubicMeter = UnitUtils.ConvertToInternalUnits(litersPerSecondCubicMeterValue, DisplayUnitType.DUT_LITERS_PER_SECOND_CUBIC_METER);
                        familyManager.Set(familyParam, convertedLitersPerSecondCubicMeter);
                    }
                    break;

                case "ElectricalIlluminance":
                    if (double.TryParse(parameterValue, out double luxValue))
                    {
                        double convertedLux = UnitUtils.ConvertToInternalUnits(luxValue, DisplayUnitType.DUT_LUX);
                        familyManager.Set(familyParam, convertedLux);
                    }
                    break;

                case "ElectricalEfficacy":
                    if (double.TryParse(parameterValue, out double lumensPerWattValue))
                    {
                        double convertedLumensPerWatt = UnitUtils.ConvertToInternalUnits(lumensPerWattValue, DisplayUnitType.DUT_LUMENS_PER_WATT);
                        familyManager.Set(familyParam, convertedLumensPerWatt);
                    }
                    break;

                case "ElectricalLuminance":
                    if (double.TryParse(parameterValue, out double candelasPerSquareMeterValue))
                    {
                        double convertedCandelasPerSquareMeter = UnitUtils.ConvertToInternalUnits(candelasPerSquareMeterValue, DisplayUnitType.DUT_CANDELAS_PER_SQUARE_METER);
                        familyManager.Set(familyParam, convertedCandelasPerSquareMeter);
                    }
                    break;

                case "ElectricalTemperature":
                case "HVACTemperature":
                case "PipingTemperature":
                    if (double.TryParse(parameterValue, out double celsiusValue))
                    {
                        double convertedCelsius = UnitUtils.ConvertToInternalUnits(celsiusValue, DisplayUnitType.DUT_CELSIUS);
                        familyManager.Set(familyParam, convertedCelsius);
                    }
                    break;    
            }
        }

        // Добавление параметров в семейство
        public void AddParametersToFamily(Document _doc, string activeFamilyName, Dictionary<string, List<string>> allParametersForAddDict)
        {
            string logFile = "ОТЧЁТ ОБ ОШИБКАХ.\n" + $"Параметры, которые не были добавлены в семейство {activeFamilyName}:\n";
            bool starusAddParametersToFamily = false;

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

                    bool isInstance = typeOrInstance.Equals("Экземпляр", StringComparison.OrdinalIgnoreCase);

                    FamilyManager familyManager = _doc.FamilyManager;
                    ExternalDefinition externalDef = parameterDefinition as ExternalDefinition;

                    FamilyParameter existingParam = familyManager.get_Parameter(externalDef);

                    if (existingParam == null)
                    {
                        FamilyParameter familyParam = familyManager.AddParameter(externalDef, grouping, isInstance);
                        starusAddParametersToFamily = true;

                        if (familyParam != null && parameterValue != "None")
                        {
                            try
                            {
                                RelationshipOfValuesWithTypesToAddToParameter(familyManager, familyParam, parameterValue, parameterValueDataType);
                                starusAddParametersToFamily = true;
                            }
                            catch (Exception ex)
                            {
                                logFile += $"Error: {generalParametersFileLink}: {parameterGroup} - {parameterName}. Группирование: {paramDetails[4]} . Экземпляр: {isInstance}. (!) ОШИБКА ДОБАВЛЕНИЯ ЗНАЧЕНИЯ: {parameterValue}\n";
                                paramDetails[5] = "!ОШИБКА";
                            }
                        }
                    }
                    else if (existingParam != null && parameterValue != "None")
                    {
                        try
                        {
                            RelationshipOfValuesWithTypesToAddToParameter(familyManager, existingParam, parameterValue, parameterValueDataType);
                            starusAddParametersToFamily = true;
                        }
                        catch (Exception ex)
                        {
                            logFile += $"Error: {generalParametersFileLink}: {parameterGroup} - {parameterName}. Группирование: {paramDetails[4]} . Экземпляр: {isInstance}. (!) ОШИБКА ОБНОВЛЕНИЯ ЗНАЧЕНИЯ: {parameterValue}\n";
                            paramDetails[5] = "!ОШИБКА";
                        }

                    }
                }

                trans.Commit();

                // Отчёт о результате выполнения в виде диалоговых окон + обновлкение полей с неисправными параметрами
                if (logFile.Contains("Error"))
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
                                            }
                                            break;
                                        }
                                    }
                                }

                                index++;
                            }
                        }
                    }

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

                            System.IO.File.WriteAllText(filePath, logFile);
                        }
                    }
                }
                else if (starusAddParametersToFamily)
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

        // Проверка заполнености поля со значение мпараметра внутри #uniqueParameterField
        public bool CheckingFillingValidDataInParamValue()
        {
            bool validDataValue = true;

            var uniquePanels = SP_allPanelParamsFields.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null && sp.Tag.ToString() == "uniqueParameterField");

            foreach (var panel in uniquePanels)
            {
                var textBox = panel.Children.OfType<System.Windows.Controls.TextBox>().FirstOrDefault();

                if (textBox != null && textBox.Tag != null && textBox.Tag.ToString() == "invalid")
                {
                    validDataValue = false;
                }
            }

                return validDataValue;
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

        // Очистка неправильно заполненных ComboBox "Параметры"
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

        //// XAML. Открытие файла ФОПа при помощи кнопки
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckingFillingAllComboBoxes())
            {
                ClearingIncorrectlyFilledFieldsParams();

                System.Windows.Forms.MessageBox.Show("Не все поля заполнены. Чтобы выбрать новый ФОП, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
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

                ClearingIncorrectlyFilledFieldsParams();
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

                if (TB_filePath.Text != null && CB_paramsGroup.SelectedItem != null && CB_paramsName.SelectedItem != null)
                {
                    try
                    {
                        revitApp.SharedParametersFilename = TB_filePath.Text;

                        DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                        DefinitionGroup defGroup = defFile.Groups.get_Item(CB_paramsGroup.SelectedItem.ToString());
                        ExternalDefinition def = defGroup.Definitions.get_Item(CB_paramsName.SelectedItem.ToString()) as ExternalDefinition;

                        ParameterType paramType = def.ParameterType;

                        CB_paramsName.Tag = paramType;
                        TB_paramValue.IsEnabled = true;                       
                        TB_paramValue.Tag = "nonestatus";
                        TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {paramType.ToString()})";
                        TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                    catch (Exception ex)
                    {
                        TB_paramValue.Text = "Не удалось прочитать параметр. Тип данных: ОШИБКА";
                        CB_paramsName.Tag = "ОШИБКА";
                        TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
                    }
                }
            }          
        }

        //// XAML.Оригинальный ComboBox "Параметры": потеря фокуса
        private void ParamsName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (CB_paramsName.SelectedItem == null)
            {
                CB_paramsName.Text = "";
                TB_paramValue.IsEnabled = false;
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));                
                TB_paramValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
            }
        }

        //// XAML.Оригинальный TextBox "Значение параметра": получение фокуса
        private void DataVerification_GotFocus(object sender, RoutedEventArgs e)
        {
            String textInField = TB_paramValue.Text;

            if (textInField.Contains("При необходимости, вы можете указать значение параметра") 
                || textInField.Contains("Необходимо указать:"))               
            {
                TB_paramValue.Clear();
            }
        }

        //// XAML.Оригинальный TextBox "Значение параметра": потеря фокуса
        private void DataVerification_LostFocus(object sender, RoutedEventArgs e)
        {         
            if (string.IsNullOrEmpty(TB_paramValue.Text))
            {
                TB_paramValue.Tag = "nonestatus";
                TB_paramValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {CB_paramsName.Tag})";
                TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
            } 
            else 
            {
                ParameterType paramType = (ParameterType)CB_paramsName.Tag;

                if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramType) == "red")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                    TB_paramValue.Tag = "invalid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramType) == "green")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName,TB_paramValue, paramType) == "blue")
                {
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                    TB_paramValue.Tag = "valid";
                }
                else if (CheckingValueOfAParameter(CB_paramsName, TB_paramValue, paramType) == "yellow")
                {
                    CB_paramsName.Text = "";
                    TB_paramValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый                   
                    TB_paramValue.IsEnabled = false;
                    TB_paramValue.Tag = "nonestatus";
                    TB_paramValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
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

            foreach (var key in groupAndParametersFromSPFDict.Keys)
            {
                cbParamsGroup.Items.Add(key);
            }

            allParamNameList = CreateallParamNameList(groupAndParametersFromSPFDict);

            foreach (var param in allParamNameList)
            {
                cbParamsName.Items.Add(param);
            }
                  
            System.Windows.Controls.ComboBox cbTypeInstance = new System.Windows.Controls.ComboBox
            {
                Width = 105,
                Height = 25,
                VerticalAlignment = VerticalAlignment.Top,
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
                VerticalAlignment = VerticalAlignment.Top,
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
                VerticalAlignment = VerticalAlignment.Top,
                Content = "X",
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            removeButton.Click += (s, ev) =>
            {
                SP_allPanelParamsFields.Children.Remove(newPanel);
            };

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
                if (cbParamsGroup.SelectedItem != null && cbParamsGroup.SelectedIndex != -1)
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

                    ClearingIncorrectlyFilledFieldsParams();
                }
            };

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

            cbParamsName.DropDownOpened += (s, ev) =>
            {
                if (cbParamsName.SelectedItem != null)
                {
                    string comboBoxValue = cbParamsName.SelectedItem.ToString();
                    cbParamsName.SelectedItem = null;
                    cbParamsName.Text = comboBoxValue;

                    var collectionViewOriginal = CollectionViewSource.GetDefaultView(cbParamsName.Items);

                    if (collectionViewOriginal != null)
                    {
                        collectionViewOriginal.Filter = null;
                        collectionViewOriginal.Refresh();
                    }
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

                    if (TB_filePath.Text != null && cbParamsGroup.SelectedItem != null && cbParamsName.SelectedItem != null)
                    {
                        try
                        {
                            revitApp.SharedParametersFilename = TB_filePath.Text;

                            DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                            DefinitionGroup defGroup = defFile.Groups.get_Item(cbParamsGroup.SelectedItem.ToString());
                            ExternalDefinition def = defGroup.Definitions.get_Item(cbParamsName.SelectedItem.ToString()) as ExternalDefinition;

                            ParameterType paramType = def.ParameterType;
                            cbParamsName.Tag = paramType;
                            tbParamValue.IsEnabled = true;
                            tbParamValue.Tag = "nonestatus";
                            tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {paramType.ToString()})";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                        }
                        catch (Exception ex)
                        {
                            tbParamValue.Text = "Не удалось прочитать параметр. Тип данных: ОШИБКА";
                            cbParamsName.Tag = "ОШИБКА";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101));
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

            tbParamValue.GotFocus += (s, ev) =>
            {
                String textInField = tbParamValue.Text;

                if (textInField.Contains("При необходимости, вы можете указать значение параметра")
                    || textInField.Contains("Необходимо указать:"))
                {
                    tbParamValue.Clear();
                }
            };

            tbParamValue.LostFocus += (s, ev) =>
            {
                if (string.IsNullOrEmpty(tbParamValue.Text))
                {
                    tbParamValue.Tag = "nonestatus";
                    tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {cbParamsName.Tag})";
                    tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                }
                else
                {
                    ParameterType paramType = (ParameterType)cbParamsName.Tag;

                    if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "red")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                        tbParamValue.Tag = "invalid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "green")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "blue")
                    {
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                        tbParamValue.Tag = "valid";
                    }
                    else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "yellow")
                    {
                        cbParamsName.Text = "";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый
                        tbParamValue.Tag = "nonestatus";
                        tbParamValue.IsEnabled = false;
                        tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
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

        //// XAML. Добавление новой панели параметров uniqueParameterField из JSON
        private void AddPanelParamFieldsJson(Dictionary<string, List<string>> allParamInInterfaceFromJsonDict)
        {
            // Уникальный ФОП для каждого uniqueParameterField параметра
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

                    if (defFile == null)
                    {
                        System.Windows.Forms.MessageBox.Show($"ФОП ``{SPFPath}``\n" +
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
                    System.Windows.Forms.MessageBox.Show($"ФОП ``{SPFPath}``\n" +
                        "не найден или неисправен. Работа плагина остановлена", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    break;
                }

                // Создание новой StackPanel #uniqueParameterField 
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
                    IsEnabled = false,
                    Width = 270,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };

                cbParamsGroup.SelectedItem = allParamInInterfaceFromJsonValues[1];

                System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
                {
                    IsEnabled = false,
                    Width = 490,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
                    Padding = new Thickness(8, 4, 0, 0),
                    Foreground = Brushes.Gray
                };
               
                cbParamsName.SelectedItem = allParamInInterfaceFromJsonValues[2];
                cbParamsName.Tag = allParamInInterfaceFromJsonValues[6];

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
                    SelectedIndex = 0
                };

                foreach (string key in CreateTypeInstanceList())
                {
                    cbTypeInstance.Items.Add(key);
                }

                foreach (var key in groupAndParametersFromSPFDict.Keys)
                {
                    cbParamsGroup.Items.Add(key);
                }

                System.Windows.Controls.ComboBox cbGrouping = new System.Windows.Controls.ComboBox
                {
                    Width = 340,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Top,
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
                    VerticalAlignment = VerticalAlignment.Top,
                    Content = "X",
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                    Foreground = new SolidColorBrush(Colors.White)
                };

                removeButton.Click += (s, ev) =>
                {
                    SP_allPanelParamsFields.Children.Remove(newPanel);
                };

                System.Windows.Controls.TextBox tbParamValue = new System.Windows.Controls.TextBox
                {
                    Width = 1245,
                    Height = 25,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Margin = new Thickness(-1250, 0, 0, 0),
                    Padding = new Thickness(15, 3, 0, 0),
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213))
                };

                tbParamValue.Text = allParamInInterfaceFromJsonValues[5];

                if (!groupAndParametersFromSPFDict.ContainsKey(allParamInInterfaceFromJsonValues[1]))
                {
                    System.Windows.Forms.MessageBox.Show($"Параметр ``{allParamInInterfaceFromJsonValues[1]}`` не найден в ФОП\n" +
                        $"(``{SPFPath})\n" +
                        "Работа плагина остановлена.\n", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    break;
                }
                              
                if (CreateTypeInstanceList().Contains(allParamInInterfaceFromJsonValues[3]))
                {
                    cbTypeInstance.SelectedItem = allParamInInterfaceFromJsonValues[3];
                } else 
                {
                    cbTypeInstance.SelectedIndex = -1;
                }             

                if (CreateGroupingDictionary().ContainsKey(allParamInInterfaceFromJsonValues[4]))
                {
                    cbGrouping.SelectedItem = allParamInInterfaceFromJsonValues[4];
                }
                else
                {
                    cbGrouping.SelectedIndex = -1;
                }
              
                tbParamValue.Loaded += (s, ev) =>
                {
                    if (tbParamValue.Text == "None")
                    {
                        tbParamValue.Tag = "nonestatus";
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {allParamInInterfaceFromJsonValues[6]})";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                    else
                    {
                        ParameterType paramType = (ParameterType)Enum.Parse(typeof(ParameterType), allParamInInterfaceFromJsonValues[6]);

                        if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "red")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                            tbParamValue.Tag = "invalid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "green")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "blue")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "yellow")
                        {
                            cbParamsName.Text = "";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый
                            tbParamValue.Tag = "nonestatus";
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
                        }
                    }
                };

                tbParamValue.GotFocus += (s, ev) =>
                {
                    String textInField = tbParamValue.Text;

                    if (textInField.Contains("При необходимости, вы можете указать значение параметра")
                        || textInField.Contains("Необходимо указать:"))
                    {
                        tbParamValue.Clear();
                    }
                };

                tbParamValue.LostFocus += (s, ev) =>
                {
                    if (string.IsNullOrEmpty(tbParamValue.Text))
                    {
                        tbParamValue.Tag = "nonestatus";
                        tbParamValue.Text = $"При необходимости, вы можете указать значение параметра (тип данных: {allParamInInterfaceFromJsonValues[6]})";
                        tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213));
                    }
                    else
                    {
                        ParameterType paramType = (ParameterType)Enum.Parse(typeof(ParameterType), allParamInInterfaceFromJsonValues[6]);

                        if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "red")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(241, 101, 101)); // Красный
                            tbParamValue.Tag = "invalid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "green")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 195, 117)); // Зелёный
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "blue")
                        {
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(70, 130, 180)); // Синий
                            tbParamValue.Tag = "valid";
                        }
                        else if (CheckingValueOfAParameter(cbParamsName, tbParamValue, paramType) == "yellow")
                        {
                            cbParamsName.Text = "";
                            tbParamValue.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(251, 255, 213)); // Жёлтый
                            tbParamValue.Tag = "nonestatus";
                            tbParamValue.IsEnabled = false;
                            tbParamValue.Text = $"Выберите значение в поле ``Группа`` или ``Параметр``";
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
                ClearingIncorrectlyFilledFieldsParams();

                System.Windows.Forms.MessageBox.Show("Не все поля заполнены\n" +
                    "Чтобы добавить параметры в семейство, заполните отсутствующие данные или удалите пустые уже существующие параметры и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            } else if (!CheckingFillingValidDataInParamValue())
            {
                System.Windows.Forms.MessageBox.Show("Не все значения параметров заполнены\n" +
                    "Чтобы добавить параметры в семейство, заполните неверно заполененые значения параметров и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else
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
                ClearingIncorrectlyFilledFieldsParams();

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
            else if (!CheckingFillingValidDataInParamValue())
            {
                System.Windows.Forms.MessageBox.Show("Не все значения параметров заполнены\n" +
                    "Чтобы сохранить файл параметров, заполните неверно заполененые значения параметров и повторите попытку.", "Не все поля заполнены",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
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
        }
    }
}