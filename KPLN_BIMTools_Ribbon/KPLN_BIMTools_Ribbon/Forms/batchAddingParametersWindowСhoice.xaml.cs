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
        UIApplication uiapp;
        public string activeFamilyName;
        public bool isFamilyDocument;

        public string paramAction;
        public string jsonFileSettingPath;

        public batchAddingParametersWindowСhoice(UIApplication uiapp, string activeFamilyName)
        {        
            InitializeComponent();
            this.uiapp = uiapp;
            this.activeFamilyName = activeFamilyName;

            if (activeFamilyName == "Семейство не выбрано")
            {
                familyName.Text = activeFamilyName;
                btnAddParamsFromSPF.IsEnabled = false;
                btnAddCustomParams.IsEnabled = false;
                btnJSONForOneParameter.IsEnabled = false;
            }

            familyName.Text = activeFamilyName;
        }

        /// <summary>
        /// Функция, которая возвращает список "Тип"/"Экземпляр"
        /// </summary>
        static public List<string> CreateTypeInstanceList() => new List<string> { "Тип", "Экземпляр" };

        /// <summary>
        /// Функция, которая возвращает типы данных в зависимости от версии Revit
        /// </summary>
#if Revit2020 || Debug2020
        static public ParameterType GetParameterTypeFromString(string dataType)
        {
            switch (dataType)
            {
                // Общие
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
                // Конструкция (Несущие конструкции)
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
                // ОВК
                case "Плотность (ОВК)": return ParameterType.HVACDensity;
                case "Трение (ОВК)": return ParameterType.HVACFriction;
                case "Мощность (ОВК)": return ParameterType.HVACPower;
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
                // Электрооборудование (Электросети)
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
                // Трубопровод (Трубопроводы)
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
                // Энергия
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
        /// <summary>
        /// Получение ForgeTypeId типа данных из string
        /// </summary>
        static public ForgeTypeId GetParameterTypeFromString(string dataType)
        {
            switch (dataType)
            {
                // Общие
                case "Текст": return SpecTypeId.String.Text; 
                case "Целое": return SpecTypeId.Int.Integer;
                case "Угол": return SpecTypeId.Angle;
                case "Площадь": return SpecTypeId.Area;
                case "Стоимость на единицу площади": return SpecTypeId.CostPerArea;
                case "Расстояние": return SpecTypeId.Distance;
                case "Длина": return SpecTypeId.Length;
                case "Массовая плотность": return SpecTypeId.MassDensity;
                case "Число": return SpecTypeId.Number;
                case "Угол поворота": return SpecTypeId.Angle;
                case "Уклон (Общие)": return SpecTypeId.Slope;
                case "Скорость (Общие)": return SpecTypeId.Speed;
                case "Время": return SpecTypeId.Time;
                case "Объем (Общие)": return SpecTypeId.Volume;
                case "Денежная единица": return SpecTypeId.Currency;
                case "URL": return SpecTypeId.String.Url;
                case "Материал": return SpecTypeId.Reference.Material;
                case "Образец заливки": return SpecTypeId.String.Text;
                case "Изображение": return SpecTypeId.Reference.Image;
                case "Да/Нет": return SpecTypeId.Boolean.YesNo;
                case "Многострочный текст": return SpecTypeId.String.MultilineText;
                // Электрооборудование (Электросети)
                case "Полная установленная мощность": return SpecTypeId.ApparentPower;
                case "Полная удельная мощность": return SpecTypeId.ApparentPowerDensity;
                case "Размер кабельного лотка": return SpecTypeId.CableTraySize;
                case "Цветовая температура": return SpecTypeId.ColorTemperature;
                case "Размер короба": return SpecTypeId.ConduitSize;
                case "Норма затрат на электроэнергию": return SpecTypeId.CostRateEnergy;
                case "Норма затрат на энергопотребление": return SpecTypeId.CostRatePower;
                case "Ток": return SpecTypeId.Current;
                case "Коэффициент спроса нагрузки": return SpecTypeId.DemandFactor;
                case "Эффективность": return SpecTypeId.Efficacy;
                case "Частота (Электросети)": return SpecTypeId.ElectricalFrequency;
                case "Освещенность": return SpecTypeId.Illuminance;
                case "Яркость": return SpecTypeId.Luminance;
                case "Световой поток": return SpecTypeId.LuminousFlux;
                case "Сила света": return SpecTypeId.LuminousIntensity;
                case "Электрический потенциал": return SpecTypeId.ElectricalPotential;
                case "Мощность (Электросети)": return SpecTypeId.ElectricalPower;
                case "Удельная мощность (Электросети)": return SpecTypeId.ElectricalPowerDensity;
                case "Мощность на единицу длины": return SpecTypeId.PowerPerLength;
                case "Электрическое удельное сопротивление": return SpecTypeId.ElectricalResistivity;
                case "Температура (Электросети)": return SpecTypeId.ElectricalTemperature;
                case "Перепад температур (Электросети)": return SpecTypeId.ElectricalTemperatureDifference;
                case "Активная мощность": return SpecTypeId.Wattage;
                case "Диаметр провода": return SpecTypeId.WireDiameter;
                case "Количество полюсов": return SpecTypeId.Int.NumberOfPoles;
                case "Классификация нагрузок": return SpecTypeId.Reference.LoadClassification;              
                // Энергия
                case "Энергия (Энергия)": return SpecTypeId.HvacEnergy;
                case "Нагревающая способность на единицу площади": return SpecTypeId.HeatCapacityPerArea;
                case "Коэффициент теплопередачи": return SpecTypeId.HeatTransferCoefficient;
                case "Изотермическая влагоемкость": return SpecTypeId.IsothermalMoistureCapacity;
                case "Проницаемость": return SpecTypeId.Permeability;
                case "Удельная теплоемкость": return SpecTypeId.SpecificHeat;
                case "Удельная теплоемкость парообразования": return SpecTypeId.SpecificHeatOfVaporization;
                case "Теплопроводность": return SpecTypeId.ThermalConductivity;
                case "Коэффициент температурного градиента для влагоемкости": return SpecTypeId.ThermalGradientCoefficientForMoistureCapacity;
                case "Тепловая нагрузка": return SpecTypeId.ThermalMass;
                case "Термостойкость": return SpecTypeId.ThermalResistance;                                                     
                // ОВК
                case "Воздушный поток": return SpecTypeId.AirFlow;
                case "Плотность воздушного потока": return SpecTypeId.AirFlowDensity;
                case "Воздушный поток, разделенный на холодильную нагрузку": return SpecTypeId.AirFlowDividedByCoolingLoad;
                case "Воздушный поток, разделенный на объем": return SpecTypeId.AirFlowDividedByVolume;
                case "Угловая скорость": return SpecTypeId.AngularSpeed;
                case "Площадь, отнесенная к холодильной нагрузке": return SpecTypeId.AreaDividedByCoolingLoad;
                case "Площадь на единицу отопительной нагрузки": return SpecTypeId.AreaDividedByHeatingLoad;
                case "Холодильная нагрузка": return SpecTypeId.CoolingLoad;
                case "Холодильная нагрузка на единицу площади": return SpecTypeId.CoolingLoadDividedByArea;
                case "Холодильная нагрузка, разделенная на объем": return SpecTypeId.CoolingLoadDividedByArea;
                case "Поперечное сечение": return SpecTypeId.CrossSection;
                case "Плотность (ОВК)": return SpecTypeId.HvacDensity;
                case "Коэффициент температуропроводимости": return SpecTypeId.Diffusivity;
                case "Толщина изоляции воздуховода": return SpecTypeId.DuctInsulationThickness;
                case "Толщина внутренней изоляции воздуховода": return SpecTypeId.DuctLiningThickness;
                case "Размер воздуховода": return SpecTypeId.DuctSize;
                case "Коэффициент": return SpecTypeId.Factor;
                case "Поток на единицу мощности": return SpecTypeId.FlowPerPower;
                case "Трение (ОВК)": return SpecTypeId.HvacFriction;
                case "Теплоприток": return SpecTypeId.HeatGain;
                case "Отопительная нагрузка": return SpecTypeId.HeatingLoad;
                case "Отопительная нагрузка на единицу площади": return SpecTypeId.HeatingLoadDividedByArea;
                case "Отопительная нагрузка, разделенная на объем": return SpecTypeId.HeatingLoadDividedByVolume;
                case "Масса на единицу времени (ОВК)": return SpecTypeId.HvacMassPerTime;
                case "Мощность (ОВК)": return SpecTypeId.HvacPower;
                case "Удельная мощность (ОВК)": return SpecTypeId.HvacPowerDensity;
                case "Мощность на единицу потока": return SpecTypeId.PowerPerFlow;
                case "Давление (ОВК)": return SpecTypeId.HvacPressure;
                case "Шероховатость (ОВК)": return SpecTypeId.HvacRoughness;
                case "Уклон (ОВК)": return SpecTypeId.HvacSlope;
                case "Температура (ОВК)": return SpecTypeId.HvacTemperature;
                case "Перепад температур (ОВК)": return SpecTypeId.HvacTemperatureDifference;
                case "Скорость (ОВК)": return SpecTypeId.HvacVelocity;
                case "Динамическая вязкость (ОВК)": return SpecTypeId.HvacViscosity;                                   
                // Инфраструктура
                case "Пикетаж": return SpecTypeId.Stationing;
                case "Интенрвал пикетов": return SpecTypeId.StationingInterval;               
                // Трубопровод (Трубопроводы)
                case "Плотность (Трубопроводы)": return SpecTypeId.PipingDensity;
                case "Расход": return SpecTypeId.Flow;
                case "Трение (Трубопроводы)": return SpecTypeId.PipingFriction;
                case "Масса (Трубопроводы)": return SpecTypeId.PipingMass;
                case "Масса на единицу времени (Трубопроводы)": return SpecTypeId.PipingMassPerTime;
                case "Величина трубы": return SpecTypeId.PipeDimension;
                case "Толщина изоляции трубы": return SpecTypeId.PipeInsulationThickness;
                case "Масса на единицу длины (Трубопроводы)": return SpecTypeId.PipeMassPerUnitLength;
                case "Размер трубы": return SpecTypeId.PipeSize;
                case "Давление (Трубопроводы)": return SpecTypeId.PipingPressure;
                case "Шероховатость (Трубопроводы)": return SpecTypeId.PipingRoughness;
                case "Уклон (Трубопроводы)": return SpecTypeId.PipingSlope;
                case "Температура (Трубопровод)": return SpecTypeId.PipingTemperature;
                case "Перепад температур (Трубопроводы)": return SpecTypeId.PipingTemperatureDifference;
                case "Скорость (Трубопроводы)": return SpecTypeId.PipingVelocity;
                case "Динамическая вязкость (Трубопроводы)": return SpecTypeId.PipingViscosity;
                case "Объем (Трубопроводы)": return SpecTypeId.PipingVolume;                                   
                // Конструкция (Несущие конструкции)
                case "Ускорение": return SpecTypeId.Acceleration;
                case "Распределенная нагрузка": return SpecTypeId.AreaForce;
                case "Коэффициент упругости среды": return SpecTypeId.AreaSpringCoefficient;
                case "Диаметр стержня": return SpecTypeId.BarDiameter;
                case "Ширина трещины": return SpecTypeId.CrackWidth;
                case "Смещение/прогиб": return SpecTypeId.Displacement;
                case "Энергия (Несущие конструкции)": return SpecTypeId.Energy;
                case "Усилие": return SpecTypeId.Force;
                case "Частота (Несущие конструкции)": return SpecTypeId.StructuralFrequency;
                case "Линейный коэффициент упругости": return SpecTypeId.LineSpringCoefficient;
                case "Распределенная нагрузка по линии": return SpecTypeId.LinearForce;
                case "Линейный момент": return SpecTypeId.LinearMoment;
                case "Масса (Несущие конструкции)": return SpecTypeId.Mass;
                case "Масса на единицу площади": return SpecTypeId.MassPerUnitArea;
                case "Масса на единицу длины (Несущие конструкции)": return SpecTypeId.MassPerUnitLength;
                case "Момент": return SpecTypeId.Moment;
                case "Момент инерции": return SpecTypeId.MomentOfInertia;
                case "Период": return SpecTypeId.Period;
                case "Сосредоточенный коэффициент упругости": return SpecTypeId.PointSpringCoefficient;
                case "Пульсация": return SpecTypeId.Pulsation;
                case "Площадь армирования": return SpecTypeId.ReinforcementArea;
                case "Армирование по площади на единицу длины": return SpecTypeId.ReinforcementAreaPerUnitLength;
                case "Защитный слой армирования": return SpecTypeId.ReinforcementCover;
                case "Длина армирования": return SpecTypeId.ReinforcementLength;
                case "Интервал армирования": return SpecTypeId.ReinforcementSpacing;
                case "Объем арматуры": return SpecTypeId.ReinforcementVolume;
                case "Поворот": return SpecTypeId.Rotation;
                case "Линейный угловой коэффициент упругости": return SpecTypeId.RotationalLineSpringCoefficient;
                case "Сосредоточенный угловой коэффициент упругости": return SpecTypeId.RotationalPointSpringCoefficient;
                case "Площадь сечения": return SpecTypeId.SectionArea;
                case "Размеры сечения": return SpecTypeId.SectionDimension;
                case "Момент сопротивления сечения": return SpecTypeId.SectionModulus;
                case "Свойство сечения": return SpecTypeId.SectionProperty;
                case "Напряжение": return SpecTypeId.Stress;
                case "Площадь поверхности на единицу длины": return SpecTypeId.SurfaceAreaPerUnitLength;
                case "Коэффициент теплового расширения": return SpecTypeId.ThermalExpansionCoefficient;
                case "Удельный вес": return SpecTypeId.UnitWeight;
                case "Скорость (Несущие конструкции)": return SpecTypeId.StructuralVelocity;
                case "Постоянная перекоса": return SpecTypeId.WarpingConstant;
                case "Вес": return SpecTypeId.Weight;
                case "Вес на единицу длины": return SpecTypeId.WeightPerUnitLength;

                default: return SpecTypeId.String.Text;
            }
        }

        /// <summary>
        /// Функция предопределения базовых типов для ForgeTypeId при создании параметра. 
        /// Все числовые значения и (!) FamilyType попадают в Autodesk.Revit.DB.ForgeTypeId
        /// </summary>
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
            if (value == SpecTypeId.Reference.Image)
                return "Image";
            if (value == SpecTypeId.Reference.LoadClassification)
                return "LoadClassification";

            return value.ToString();
        }
#endif

        /// <summary>
        /// Создание Dictionary с параметрами группирования для ComboBox "Параметры группирования"
        /// </summary>
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
                { "Электросети", BuiltInParameterGroup.PG_ELECTRICAL },
                { "Электросети - Нагрузки", BuiltInParameterGroup.PG_ELECTRICAL_LOADS },
                { "Электросети - Освещение", BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING },
                { "Электросети - Создание цепей", BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING },
#if Revit2023 || Debug2023
                { "R23. Анализ электросетей", BuiltInParameterGroup.INVALID },
                { "R23. Силы жизнеобеспечения", BuiltInParameterGroup.INVALID },
                { "R23. Электротехника", BuiltInParameterGroup.INVALID },
#endif
            };

            return groupingDict;
        }

#if Revit2023 || Debug2023
        /// <summary>
        /// Получение ForgeTypeId группирования из string
        /// </summary>
        static public ForgeTypeId GetParameterGroupFromString(string group)
        {
            switch (group)
            {
                case "Аналитическая модель": return GroupTypeId.AnalyticalModel;
                case "Видимость": return GroupTypeId.Visibility;
                case "Второстепенный конец": return GroupTypeId.SecondaryEnd;
                case "Выравнивание аналитической модели": return GroupTypeId.AnalyticalAlignment;
                case "Геометрия разделения": return GroupTypeId.DivisionGeometry;
                case "Графика": return GroupTypeId.Graphics;
                case "Данные": return GroupTypeId.Data;
                case "Зависимости": return GroupTypeId.Constraints;
                case "Идентификация": return GroupTypeId.IdentityData;
                case "Материалы и отделка": return GroupTypeId.Materials;
                case "Механизмы": return GroupTypeId.Mechanical;
                case "Механизмы - Нагрузки": return GroupTypeId.MechanicalLoads;
                case "Механизмы - Расход": return GroupTypeId.MechanicalAirflow;
                case "Моменты": return GroupTypeId.Moments;
                case "Набор": return GroupTypeId.CouplerArray;
                case "Набор арматурных стержней": return GroupTypeId.RebarArray;
                case "Несущие конструкции": return GroupTypeId.Structural;
                case "Общая легенда": return GroupTypeId.OverallLegend;
                case "Общие": return GroupTypeId.General;
                case "Основной конец": return GroupTypeId.PrimaryEnd;
                case "Параметры IFC": return GroupTypeId.Ifc;
                case "Прочее": return new ForgeTypeId(string.Empty);
                case "Размеры": return GroupTypeId.Geometry;
                case "Расчет несущих конструкций": return GroupTypeId.StructuralAnalysis;
                case "Расчет энергопотребления": return GroupTypeId.EnergyAnalysis;
                case "Редактирование формы перекрытия": return GroupTypeId.SlabShapeEdit;
                case "Результат анализа": return GroupTypeId.AnalysisResults;
                case "Сантехника": return GroupTypeId.Plumbing;
                case "Свойства модели": return GroupTypeId.AdskModelProperties;
                case "Свойства экологически чистого здания": return GroupTypeId.GreenBuilding;
                case "Сегменты и соединительные детали": return GroupTypeId.SegmentsFittings;
                case "Силы": return GroupTypeId.Forces;
                case "Система пожаротушения": return GroupTypeId.FireProtection;
                case "Слои": return GroupTypeId.RebarSystemLayers;
                case "Снятие связей/усилия для элемента": return GroupTypeId.ReleasesMemberForces;
                case "Стадии": return GroupTypeId.Phasing;
                case "Строительство": return GroupTypeId.Construction;
                case "Текст": return GroupTypeId.Text;
                case "Фотометрические": return GroupTypeId.LightPhotometrics;
                case "Шрифт заголовков": return GroupTypeId.Title;
                case "Электросети": return GroupTypeId.Electrical;
                case "Электросети - Нагрузки": return GroupTypeId.ElectricalLoads;
                case "Электросети - Освещение": return GroupTypeId.ElectricalLighting;
                case "Электросети - Создание цепей": return GroupTypeId.ElectricalCircuiting;               
                case "R23. Анализ электросетей": return GroupTypeId.ElectricalAnalysis;
                case "R23. Силы жизнеобеспечения": return GroupTypeId.LifeSafety;
                case "R23. Электротехника": return GroupTypeId.ElectricalEngineering;               
                default: return GroupTypeId.General;
            }
        }
#endif

        /// <summary>
        /// Функция соотношенияч типа данных со значением при добавлении параметра в семейство
        /// </summary>
        public void RelationshipOfValuesWithTypesToAddToParameter(FamilyManager familyManager, FamilyParameter familyParam, String parameterValue, String parameterValueDataType)
        {
            switch (parameterValueDataType)
            {
                case "Text":
                case "Текст":
                case "MultilineText":
                case "Многострочный текст":
                case "URL":
                    familyManager.Set(familyParam, parameterValue);
                    break;

                case "Integer":
                case "Целое":
                case "YesNo":
                case "Да/Нет":
                case "NumberOfPoles":
                case "Количество полюсов":
                    if (int.TryParse(parameterValue, out int intBoolValue))
                    {
                        familyManager.Set(familyParam, intBoolValue);
                    }
                    break;

                case "Material":
                case "Материал":
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

#if Revit2020 || Debug2020
                case "Image":
                case "Изображение":
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
                case "Изображение":
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

        /// <summary>
        /// Функция проверки введённого значения в поле параметра: 
        /// "yellow" (пустой параметр), "blue" (невозможно проверить), "green" (проверка пройдена), "red" (проверка не пройдена);
        /// </summary>
        public string CheckingValueOfAParameter(System.Windows.Controls.ComboBox comboBox, System.Windows.Controls.TextBox textBox, string paramTypeName)
        {
            string textInField = textBox.Text;

            if (comboBox.SelectedItem == null) return "yellow";
            if (textBox.Text.ToString().StartsWith("=") || paramTypeName == "Image" || paramTypeName == "Изображение") return "blue";

            switch (paramTypeName)
            {
                case "Text":
                case "Текст":
                case "MultilineText":
                case "Многострочный текст":
                case "URL":
                    return "green";

                case "Integer":
                case "Целое":
                    if (textInField.Contains(",")) return "red";
                    if (int.TryParse(textInField, out int resultInt))
                    {
                        textBox.Text = resultInt.ToString();
                        return "green";
                    }
                    break;

                case "NumberOfPoles":             
                case "Количество полюсов":
                    if (int.TryParse(textInField, out int resultIntU) && resultIntU >= 1 && resultIntU <= 3)
                    {
                        textBox.Text = resultIntU.ToString();
                        return "green";
                    }
                    textBox.Text = "Необходимо указать: диапазон от 1 до 3";
                    return "red";

                case "YesNo":
                case "Да/Нет":
                    if (textInField == "0" || textInField == "1") return "green";
                    textBox.Text = "Необходимо указать: ``1`` - да; ``0`` - нет";
                    return "red";

                case "Material":
                case "Материал":
                    var materialNames = new FilteredElementCollector(uiapp.ActiveUIDocument.Document)
                                        .OfClass(typeof(Material))
                                        .Cast<Material>()
                                        .Select(m => m.Name.ToLower())
                                        .ToList();
                    return materialNames.Contains(textInField.ToLower()) ? "green" : "red";
            }

            if (double.TryParse(textInField, out double resultDouble))
            {
                return "green";
            }

            return "red";
        }


        ///////////////////////////////////////////////
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
            jsonFileSettingPath = "";

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
                    System.Windows.Forms.MessageBox.Show("Ваш JSON-файл не является файлом преднастроек или повреждён. " +
                        "Пожалуйста, выберите другой файл.", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        //// XAML. Загрузка XAML-настройки для нескольких параметров
        private void Button_MultipleLoadParam(object sender, RoutedEventArgs e)
        {
            var window = new batchAddingParametersWindowMultipleLoadParameters(uiapp);
            var revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            new System.Windows.Interop.WindowInteropHelper(window).Owner = revitHandle;
            window.ShowDialog();         
        }
    }
}
