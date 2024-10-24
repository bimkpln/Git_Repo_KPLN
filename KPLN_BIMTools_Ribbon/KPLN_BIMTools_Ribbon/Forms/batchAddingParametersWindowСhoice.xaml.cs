using Autodesk.Revit.UI;
using System.Windows;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Linq;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Controls;


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

        // Функция соотношения параметра с типом данных и установки нужного кол-ва знаков после запятой.
        // Возвращает "yellow" - параметр пуст или указан неверно;
        // Возвращает "blue" - невозможно проверить значение параметра;
        // Возвращает "green" - значение параметра прошло проверку;
        // Возвращает "red" - значение параметра не прошло проверку;
        public string CheckingValueOfAParameter(System.Windows.Controls.ComboBox comboBox, System.Windows.Controls.TextBox textBox, ParameterType paramType)
        {
            var textInField = textBox.Text;

            if (comboBox.SelectedItem == null)
            {
                return "yellow";
            }

            if (paramType == ParameterType.Image)
            {
                if (!string.IsNullOrEmpty(textInField))
                {
                    return "blue";
                }
            }

          if (paramType == ParameterType.Material)
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

            if (paramType == ParameterType.NumberOfPoles)
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
                || paramType == ParameterType.PipingTemperatureDifference || paramType == ParameterType.PipingVelocity || paramType == ParameterType.ReinforcementArea || paramType == ParameterType.ReinforcementCover || paramType == ParameterType.ReinforcementLength
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

            if (paramType == ParameterType.FamilyType || paramType == ParameterType.LoadClassification)
            {               
                return "red";
            }

            return "red";
        }


        // Функция соотношенияч типа данных со значением при добавлении параметра в семейство
        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, String parameterValue, String parameterValueDataType)
        {
            switch (parameterValueDataType)
            {
                /// 
                /// Типы c зависимостями по ID
                /// 
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

#if Revit2020 || Debug2020
                ImageType newImageTypeOld = ImageType.Create(uiapp.ActiveUIDocument.Document, imagePath);
                familyManager.Set(familyParam, newImageTypeOld.Id);
#endif

#if Revit2023 || Debug2023
                ImageTypeOptions imageTypeOptions = new ImageTypeOptions(imagePath, false, ImageTypeSource.Imported);
                ImageType newImageTypeNew = ImageType.Create(uiapp.ActiveUIDocument.Document, imageTypeOptions);
                familyManager.Set(familyParam, newImageTypeNew.Id);
#endif  

                    }
                    break;

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
                        double convertedJoules = UnitUtils.ConvertToInternalUnits(joulesValue, DisplayUnitType.DUT_JOULES);
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

                case "HVACPermeability":
                    if (double.TryParse(parameterValue, out double nanogramsPerPascalSecondSquareMeterValue))
                    {
                        familyManager.Set(familyParam, 10000); // Чтобы добавить 3 значное и меньше значение
                        double convertedNanogramsPerPascalSecondSquareMeter = UnitUtils.ConvertToInternalUnits(nanogramsPerPascalSecondSquareMeterValue, DisplayUnitType.DUT_NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER);
                        familyManager.Set(familyParam, convertedNanogramsPerPascalSecondSquareMeter);
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

                string jsonContent = File.ReadAllText(jsonFileSettingPath);
                dynamic jsonFile = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);

                if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["pathFile"] != null && item["groupParameter"] != null && item["nameParameter"] != null && item["instance"] != null && item["grouping"] != null && item["parameterValue"] != null && item["parameterValueDataType"] != null))
                {
                    var window = new batchAddingParametersWindowGeneral(uiapp, activeFamilyName, jsonFileSettingPath);
                    var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
                    window.ShowDialog();
                }
                else if (jsonFile is JArray && ((JArray)jsonFile).Any(item => ((string)item["NE"])?.StartsWith("FamilyParamAdd") == true))
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
