using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class CommandLTablesExport : IExternalCommand
    {
        /// <summary>
        /// В словаре указаны обрезанные от версии id. В разных ревитах - они отличаются
        /// </summary>
        private readonly Dictionary<string, string> _forgeTypeIdDict = new Dictionary<string, string>
        {
            {"autodesk.spec.aec.electrical:apparentPower", "ELECTRICAL_APPARENT_POWER"},
            {"autodesk.spec.aec.electrical:cableTraySize", "CABLETRAY_SIZE"},
            {"autodesk.spec.aec.electrical:colorTemperature", "COLOR_TEMPERATURE"},
            {"autodesk.spec.aec.electrical:conduitSize", "CONDUIT_SIZE"},
            {"autodesk.spec.aec.electrical:costRateEnergy", "ELECTRICAL_COST_RATE_ENERGY"},
            {"autodesk.spec.aec.electrical:costRatePower", "ELECTRICAL_COST_RATE_POWER"},
            {"autodesk.spec.aec.electrical:current", "ELECTRICAL_CURRENT"},
            {"autodesk.spec.aec.electrical:demandFactor", "ELECTRICAL_DEMAND_FACTOR"},
            {"autodesk.spec.aec.electrical:efficacy", "ELECTRICAL_EFFICACY"},
            {"autodesk.spec.aec.electrical:frequency", "ELECTRICAL_FREQUENCY"},
            {"autodesk.spec.aec.electrical:illuminance", "ELECTRICAL_ILLUMINANCE"},
            {"autodesk.spec.aec.electrical:luminance", "ELECTRICAL_LUMINANCE"},
            {"autodesk.spec.aec.electrical:luminousFlux", "ELECTRICAL_LUMINOUS_FLUX"},
            {"autodesk.spec.aec.electrical:luminousIntensity", "ELECTRICAL_LUMINOUS_INTENSITY"},
            {"autodesk.spec.aec.electrical:potential", "ELECTRICAL_POTENTIAL"},
            {"autodesk.spec.aec.electrical:power", "ELECTRICAL_POWER"},
            {"autodesk.spec.aec.electrical:powerDensity", "ELECTRICAL_POWER_DENSITY"},
            {"autodesk.spec.aec.electrical:powerPerLength", "ELECTRICAL_POWER_PER_LENGTH"},
            {"autodesk.spec.aec.electrical:resistivity", "ELECTRICAL_RESISTIVITY"},
            {"autodesk.spec.aec.electrical:temperature", "ELECTRICAL_TEMPERATURE"},
            {"autodesk.spec.aec.electrical:temperatureDifference", "ELECTRICAL_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.electrical:wattage", "ELECTRICAL_WATTAGE"},
            {"autodesk.spec.aec.electrical:wireDiameter", "WIRE_SIZE"},
            {"autodesk.spec.aec.energy:energy", "HVAC_ENERGY"},
            {"autodesk.spec.aec.energy:heatTransferCoefficient", "HVAC_COEFFICIENT_OF_HEAT_TRANSFER"},
            {"autodesk.spec.aec.energy:isothermalMoistureCapacity", "HVAC_ISOTHERMAL_MOISTURE_CAPACITY"},
            {"autodesk.spec.aec.energy:permeability", "HVAC_PERMEABILITY"},
            {"autodesk.spec.aec.energy:specificHeat", "HVAC_SPECIFIC_HEAT"},
            {"autodesk.spec.aec.energy:specificHeatOfVaporization", "HVAC_SPECIFIC_HEAT_OF_VAPORIZATION"},
            {"autodesk.spec.aec.energy:thermalConductivity", "HVAC_THERMAL_CONDUCTIVITY"},
            {"autodesk.spec.aec.energy:thermalGradientCoefficientForMoistureCapacity", "HVAC_THERMAL_GRADIENT_COEFFICIENT_FOR_MOISTURE_CAPACITY"},
            {"autodesk.spec.aec.energy:thermalMass", "HVAC_THERMAL_MASS"},
            {"autodesk.spec.aec.energy:thermalResistance", "HVAC_THERMAL_RESISTANCE"},
            {"autodesk.spec.aec.hvac:airFlow", "HVAC_AIR_FLOW"},
            {"autodesk.spec.aec.hvac:airFlowDensity", "HVAC_AIRFLOW_DENSITY"},
            {"autodesk.spec.aec.hvac:airFlowDividedByCoolingLoad", "HVAC_AIRFLOW_DIVIDED_BY_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:airFlowDividedByVolume", "HVAC_AIRFLOW_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:angularSpeed", "HVAC_ANGULAR_SPEED"},
            {"autodesk.spec.aec.hvac:areaDividedByCoolingLoad", "HVAC_AREA_DIVIDED_BY_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:areaDividedByHeatingLoad", "HVAC_AREA_DIVIDED_BY_HEATING_LOAD"},
            {"autodesk.spec.aec.hvac:coolingLoad", "HVAC_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:coolingLoadDividedByArea", "HVAC_COOLING_LOAD_DIVIDED_BY_AREA"},
            {"autodesk.spec.aec.hvac:coolingLoadDividedByVolume", "HVAC_COOLING_LOAD_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:crossSection", "HVAC_CROSS_SECTION"},
            {"autodesk.spec.aec.hvac:density", "HVAC_DENSITY"},
            {"autodesk.spec.aec.hvac:diffusivity", "HVAC_DIFFUSIVITY"},
            {"autodesk.spec.aec.hvac:ductInsulationThickness", "HVAC_DUCT_INSULATION_THICKNESS"},
            {"autodesk.spec.aec.hvac:ductLiningThickness", "HVAC_DUCT_LINING_THICKNESS"},
            {"autodesk.spec.aec.hvac:ductSize", "HVAC_DUCT_SIZE"},
            {"autodesk.spec.aec.hvac:factor", "HVAC_FACTOR"},
            {"autodesk.spec.aec.hvac:flowPerPower", "HVAC_FLOW_PER_POWER"},
            {"autodesk.spec.aec.hvac:friction", "HVAC_FRICTION"},
            {"autodesk.spec.aec.hvac:heatGain", "HVAC_HEAT_GAIN"},
            {"autodesk.spec.aec.hvac:heatingLoad", "HVAC_HEATING_LOAD"},
            {"autodesk.spec.aec.hvac:heatingLoadDividedByArea", "HVAC_HEATING_LOAD_DIVIDED_BY_AREA"},
            {"autodesk.spec.aec.hvac:heatingLoadDividedByVolume", "HVAC_HEATING_LOAD_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:massPerTime", "HVAC_MASS_PER_TIME"},
            {"autodesk.spec.aec.hvac:power", "HVAC_POWER"},
            {"autodesk.spec.aec.hvac:powerDensity", "HVAC_POWER_DENSITY"},
            {"autodesk.spec.aec.hvac:powerPerFlow", "HVAC_POWER_PER_FLOW"},
            {"autodesk.spec.aec.hvac:pressure", "HVAC_PRESSURE"},
            {"autodesk.spec.aec.hvac:roughness", "HVAC_ROUGHNESS"},
            {"autodesk.spec.aec.hvac:slope", "HVAC_SLOPE"},
            {"autodesk.spec.aec.hvac:temperature", "HVAC_TEMPERATURE"},
            {"autodesk.spec.aec.hvac:temperatureDifference", "HVAC_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.hvac:velocity", "HVAC_VELOCITY"},
            {"autodesk.spec.aec.hvac:viscosity", "HVAC_VISCOSITY"},
            {"autodesk.spec.aec.infrastructure:stationing", "STATIONING"},
            {"autodesk.spec.aec.infrastructure:stationingInterval", "STATIONING_INTERVAL"},
            {"autodesk.spec.aec.piping:density", "PIPING_DENSITY"},
            {"autodesk.spec.aec.piping:flow", "PIPING_FLOW"},
            {"autodesk.spec.aec.piping:friction", "PIPING_FRICTION"},
            {"autodesk.spec.aec.piping:mass", "PIPE_MASS"},
            {"autodesk.spec.aec.piping:massPerTime", "PIPING_MASS_PER_TIME"},
            {"autodesk.spec.aec.piping:pipeDimension", "PIPE_DIMENSION"},
            {"autodesk.spec.aec.piping:pipeInsulationThickness", "PIPE_INSUlATION_THICKNESS"},
            {"autodesk.spec.aec.piping:pipeMassPerUnitLength", "PIPE_MASS_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.piping:pipeSize", "PIPE_SIZE"},
            {"autodesk.spec.aec.piping:pressure", "PIPING_PRESSURE"},
            {"autodesk.spec.aec.piping:roughness", "PIPING_ROUGHNESS"},
            {"autodesk.spec.aec.piping:slope", "PIPING_SLOPE"},
            {"autodesk.spec.aec.piping:temperature", "PIPING_TEMPERATURE"},
            {"autodesk.spec.aec.piping:temperatureDifference", "PIPING_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.piping:velocity", "PIPING_VELOCITY"},
            {"autodesk.spec.aec.piping:viscosity", "PIPING_VISCOSITY"},
            {"autodesk.spec.aec.piping:volume", "PIPING_VOLUME"},
            {"autodesk.spec.aec.structural:acceleration", "ACCELERATION"},
            {"autodesk.spec.aec.structural:areaForce", "AREA_FORCE"},
            {"autodesk.spec.aec.structural:areaForceScale", "AREA_FORCE_SCALE"},
            {"autodesk.spec.aec.structural:areaSpringCoefficient", "AREA_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:barDiameter", "BAR_DIAMETER"},
            {"autodesk.spec.aec.structural:crackWidth", "CRACK_WIDTH"},
            {"autodesk.spec.aec.structural:displacement", "DISPLACEMENT/DEFLECTION"},
            {"autodesk.spec.aec.structural:energy", "ENERGY"},
            {"autodesk.spec.aec.structural:force", "FORCE"},
            {"autodesk.spec.aec.structural:forceScale", "FORCE_SCALE"},
            {"autodesk.spec.aec.structural:frequency", "STRUCTURAL_FREQUENCY"},
            {"autodesk.spec.aec.structural:lineSpringCoefficient", "LINEAR_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:linearForce", "LINEAR_FORCE"},
            {"autodesk.spec.aec.structural:linearForceScale", "LINEAR_FORCE_SCALE"},
            {"autodesk.spec.aec.structural:linearMoment", "LINEAR_MOMENT"},
            {"autodesk.spec.aec.structural:linearMomentScale", "LINEAR_MOMENT_SCALE"},
            {"autodesk.spec.aec.structural:mass", "MASS"},
            {"autodesk.spec.aec.structural:massPerUnitArea", "MASS_PER_UNIT_AREA"},
            {"autodesk.spec.aec.structural:massPerUnitLength", "MASS_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.structural:moment", "MOMENT"},
            {"autodesk.spec.aec.structural:momentOfInertia", "MOMENT_OF_INERTIA"},
            {"autodesk.spec.aec.structural:momentScale", "MOMENT_SCALE"},
            {"autodesk.spec.aec.structural:period", "PERIOD"},
            {"autodesk.spec.aec.structural:pointSpringCoefficient", "POINT_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:pulsation", "PULSATION"},
            {"autodesk.spec.aec.structural:reinforcementArea", "REINFORCEMENT_AREA"},
            {"autodesk.spec.aec.structural:reinforcementAreaPerUnitLength", "REINFORCEMENT_AREA_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.structural:reinforcementCover", "REINFORCEMENT_COVER"},
            {"autodesk.spec.aec.structural:reinforcementLength", "REINFORCEMENT_LENGTH"},
            {"autodesk.spec.aec.structural:reinforcementSpacing", "REINFORCEMENT_SPACING"},
            {"autodesk.spec.aec.structural:reinforcementVolume", "REINFORCEMENT_VOLUME"},
            {"autodesk.spec.aec.structural:rotation", "ROTATION"},
            {"autodesk.spec.aec.structural:rotationalLineSpringCoefficient", "ROTATIONAL_LINEAR_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:rotationalPointSpringCoefficient", "ROTATIONAL_POINT_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:sectionArea", "SECTION_AREA"},
            {"autodesk.spec.aec.structural:sectionDimension", "SECTION_DIMENSION"},
            {"autodesk.spec.aec.structural:sectionModulus", "SECTION_MODULUS"},
            {"autodesk.spec.aec.structural:sectionProperty", "SECTION_PROPERTY"},
            {"autodesk.spec.aec.structural:stress", "STRESS"},
            {"autodesk.spec.aec.structural:surfaceAreaPerUnitLength", "SURFACE_AREA"},
            {"autodesk.spec.aec.structural:thermalExpansionCoefficient", "THERMAL_EXPANSION_COEFFICIENT"},
            {"autodesk.spec.aec.structural:unitWeight", "UNIT_WEIGHT"},
            {"autodesk.spec.aec.structural:velocity", "VELOCITY"},
            {"autodesk.spec.aec.structural:warpingConstant", "WARPING_CONSTANT"},
            {"autodesk.spec.aec.structural:weight", "WEIGHT"},
            {"autodesk.spec.aec.structural:weightPerUnitLength", "WEIGHT_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec:angle", "ANGLE"},
            {"autodesk.spec.aec:area", "AREA"},
            {"autodesk.spec.aec:costPerArea", "COST_PER_AREA"},
            {"autodesk.spec.aec:decimalSheetLength", "PEN SIZE"},
            {"autodesk.spec.aec:distance", "DISTANCE"},
            {"autodesk.spec.aec:length", "LENGTH"},
            {"autodesk.spec.aec:massDensity", "MASS_DENSITY"},
            {"autodesk.spec.aec:number", "NUMBER"},
            {"autodesk.spec.aec:rotationAngle", "ROTATION_ANGLE"},
            {"autodesk.spec.aec:sheetLength", "PAPER LENGTH"},
            {"autodesk.spec.aec:siteAngle", "SITE_ANGLE"},
            {"autodesk.spec.aec:slope", "SLOPE"},
            {"autodesk.spec.aec:speed", "SPEED"},
            {"autodesk.spec.aec:time", "TIMEINTERVAL"},
            {"autodesk.spec.aec:volume", "VOLUME"},
            {"autodesk.spec.measurable:currency", "CURRENCY"}
        };

        /// <summary>
        /// В словаре указаны обрезанные от версии id. В разных ревитах - они отличаются
        /// </summary>
        private readonly Dictionary<string, string> _unitTypeIdDict = new Dictionary<string, string>
        {
            {"autodesk.unit.unit:1ToRatio", "RATIO1"},
            {"autodesk.unit.unit:acres", "ACRES"},
            {"autodesk.unit.unit:amperes", "AMPERES"},
            {"autodesk.unit.unit:atmospheres", "ATMOSPHERES"},
            {"autodesk.unit.unit:bars", "BARS"},
            {"autodesk.unit.unit:britishThermalUnits", "BRITISH_THERMAL_UNITS"},
            {"autodesk.unit.unit:britishThermalUnitsPerDegreeFahrenheit", "BRITISH_THERMAL_UNITS_PER_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHour", "BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourCubicFoot", "BRITISH_THERMAL_UNITS_PER_HOUR_CUBIC_FOOT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourFootDegreeFahrenheit", "BRITISH_THERMAL_UNITS_PER_HOUR_FOOT_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourSquareFoot", "BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourSquareFootDegreeFahrenheit", "BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerPound", "BRITISH_THERMAL_UNITS_PER_POUND"},
            {"autodesk.unit.unit:britishThermalUnitsPerPoundDegreeFahrenheit", "BRITISH_THERMAL_UNITS_PER_POUND_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerSecond", "BRITISH_THERMAL_UNITS_PER_SECOND"},
            {"autodesk.unit.unit:calories", "CALORIES"},
            {"autodesk.unit.unit:caloriesPerSecond", "CALORIES_PER_SECOND"},
            {"autodesk.unit.unit:candelas", "CANDELAS"},
            {"autodesk.unit.unit:candelasPerSquareFoot", "CANDELAS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:candelasPerSquareMeter", "CANDELAS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:celsius", "CELSIUS"},
            {"autodesk.unit.unit:celsiusInterval", "CELSIUS_INTERVAL"},
            {"autodesk.unit.unit:centimeters", "CENTIMETERS"},
            {"autodesk.unit.unit:centimetersPerMinute", "CENTIMETERS_PER_MINUTE"},
            {"autodesk.unit.unit:centimetersToTheFourthPower", "CENTIMETERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:centimetersToTheSixthPower", "CENTIMETERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:centipoises", "CENTIPOISES"},
            {"autodesk.unit.unit:cubicCentimeters", "CUBIC_CENTIMETERS"},
            {"autodesk.unit.unit:cubicFeet", "CUBIC_FEET"},
            {"autodesk.unit.unit:cubicFeetPerHour", "CUBIC_FEET_PER_HOUR"},
            {"autodesk.unit.unit:cubicFeetPerKip", "CUBIC_FEET_PER_KIP"},
            {"autodesk.unit.unit:cubicFeetPerMinute", "CUBIC_FEET_PER_MINUTE"},
            {"autodesk.unit.unit:cubicFeetPerMinuteCubicFoot", "CUBIC_FEET_PER_MINUTE_CUBIC_FOOT"},
            {"autodesk.unit.unit:cubicFeetPerMinutePerBritishThermalUnitPerHour", "CUBIC_FEET_PER_MINUTE_PER_BRITISH_THERMAL_UNIT_PER_HOUR"},
            {"autodesk.unit.unit:cubicFeetPerMinuteSquareFoot", "CUBIC_FEET_PER_MINUTE_SQUARE_FOOT"},
            {"autodesk.unit.unit:cubicFeetPerMinuteTonOfRefrigeration", "CUBIC_FEET_PER_MINUTE_TON_OF_REFRIGERATION"},
            {"autodesk.unit.unit:cubicFeetPerPoundMass", "CUBIC_FEET_PER_POUND_MASS"},
            {"autodesk.unit.unit:cubicInches", "CUBIC_INCHES"},
            {"autodesk.unit.unit:cubicMeters", "CUBIC_METERS"},
            {"autodesk.unit.unit:cubicMetersPerHour", "CUBIC_METERS_PER_HOUR"},
            {"autodesk.unit.unit:cubicMetersPerHourCubicMeter", "CUBIC_METERS_PER_HOUR_CUBIC_METER"},
            {"autodesk.unit.unit:cubicMetersPerHourSquareMeter", "CUBIC_METERS_PER_HOUR_SQUARE_METER"},
            {"autodesk.unit.unit:cubicMetersPerKilogram", "CUBIC_METERS_PER_KILOGRAM"},
            {"autodesk.unit.unit:cubicMetersPerKilonewton", "CUBIC_METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:cubicMetersPerSecond", "CUBIC_METERS_PER_SECOND"},
            {"autodesk.unit.unit:cubicMetersPerWattSecond", "CUBIC_METERS_PER_WATT_SECOND"},
            {"autodesk.unit.unit:cubicMillimeters", "CUBIC_MILLIMETERS"},
            {"autodesk.unit.unit:cubicYards", "CUBIC_YARDS"},
            {"autodesk.unit.unit:currency", "CURRENCY"},
            {"autodesk.unit.unit:currencyPerBritishThermalUnit", "COST_PER_BRITISH_THERMAL_UNIT"},
            {"autodesk.unit.unit:currencyPerBritishThermalUnitPerHour", "COST_PER_BRITISH_THERMAL_UNIT_PER_HOUR"},
            {"autodesk.unit.unit:currencyPerSquareFoot", "COST_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:currencyPerSquareMeter", "COST_PER_SQUARE_METER"},
            {"autodesk.unit.unit:currencyPerWatt", "COST_PER_WATT"},
            {"autodesk.unit.unit:currencyPerWattHour", "COST_PER_WATT_HOUR"},
            {"autodesk.unit.unit:cyclesPerSecond", "CYCLES_PER_SECOND"},
            {"autodesk.unit.unit:decimeters", "DECIMETERS"},
            {"autodesk.unit.unit:degrees", "DEGREES"},
            {"autodesk.unit.unit:dekanewtonMeters", "DEKANEWTON_METERS"},
            {"autodesk.unit.unit:dekanewtonMetersPerMeter", "DEKANEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:dekanewtons", "DEKANEWTONS"},
            {"autodesk.unit.unit:dekanewtonsPerMeter", "DEKANEWTONS_PER_METER"},
            {"autodesk.unit.unit:dekanewtonsPerSquareMeter", "DEKANEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:fahrenheit", "FAHRENHEIT"},
            {"autodesk.unit.unit:fahrenheitInterval", "FAHRENHEIT_INTERVAL"},
            {"autodesk.unit.unit:feet", "FEET"},
            {"autodesk.unit.unit:feetOfWater39.2DegreesFahrenheit", "FEET_OF_WATER"},
            {"autodesk.unit.unit:feetOfWater39.2DegreesFahrenheitPer100Feet", "FEET_OF_WATER_PER_100FT"},
            {"autodesk.unit.unit:feetPerKip", "FEET_PER_KIP"},
            {"autodesk.unit.unit:feetPerMinute", "FEET_PER_MINUTE"},
            {"autodesk.unit.unit:feetPerSecond", "FEET_PER_SECOND"},
            {"autodesk.unit.unit:feetPerSecondSquared", "FEET_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:feetToTheFourthPower", "FEET_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:feetToTheSixthPower", "FEET_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:fixed", "FIXED"},
            {"autodesk.unit.unit:footcandles", "FOOTCANDLES"},
            {"autodesk.unit.unit:footlamberts", "FOOTLAMBERTS"},
            {"autodesk.unit.unit:general", "GENERAL"},
            {"autodesk.unit.unit:gradians", "GRADIANS"},
            {"autodesk.unit.unit:grains", "GRAINS"},
            {"autodesk.unit.unit:grainsPerHourSquareFootInchMercury", "GRAINS_PER_HOUR_SQUARE_FOOT_INCH_MERCURY"},
            {"autodesk.unit.unit:grams", "GRAMS"},
            {"autodesk.unit.unit:hectares", "HECTARES"},
            {"autodesk.unit.unit:hectometers", "HECTOMETERS"},
            {"autodesk.unit.unit:hertz", "HERTZ"},
            {"autodesk.unit.unit:horsepower", "HORSEPOWER"},
            {"autodesk.unit.unit:hourSquareFootDegreesFahrenheitPerBritishThermalUnit", "HOUR_SQUARE_FOOT_DEGREES_FAHRENHEIT_PER_BRITISH_THERMAL_UNIT"},
            {"autodesk.unit.unit:hours", "HOURS"},
            {"autodesk.unit.unit:inches", "INCHES"},
            {"autodesk.unit.unit:inchesOfMercury32DegreesFahrenheit", "INCHES_OF_MERCURY"},
            {"autodesk.unit.unit:inchesOfWater60DegreesFahrenheit", "INCHES_OF_WATER"},
            {"autodesk.unit.unit:inchesOfWater60DegreesFahrenheitPer100Feet", "INCHES_OF_WATER_PER_100FT"},
            {"autodesk.unit.unit:inchesPerSecond", "INCHES_PER_SECOND"},
            {"autodesk.unit.unit:inchesPerSecondSquared", "INCHES_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:inchesToTheFourthPower", "INCHES_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:inchesToTheSixthPower", "INCHES_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:inverseDegreesCelsius", "INVERSE_DEGREES_CELSIUS"},
            {"autodesk.unit.unit:inverseDegreesFahrenheit", "INVERSE_DEGREES_FAHRENHEIT"},
            {"autodesk.unit.unit:inverseKilonewtons", "INVERSE_KILONEWTONS"},
            {"autodesk.unit.unit:inverseKips", "INVERSE_KIPS"},
            {"autodesk.unit.unit:joules", "JOULES"},
            {"autodesk.unit.unit:joulesPerGram", "JOULES_PER_GRAM"},
            {"autodesk.unit.unit:joulesPerGramDegreeCelsius", "JOULES_PER_GRAM_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:joulesPerKelvin", "JOULES_PER_KELVIN"},
            {"autodesk.unit.unit:joulesPerKilogram", "JOULES_PER_KILOGRAM"},
            {"autodesk.unit.unit:joulesPerKilogramDegreeCelsius", "JOULES_PER_KILOGRAM_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:kelvin", "KELVIN"},
            {"autodesk.unit.unit:kelvinInterval", "KELVIN_INTERVAL"},
            {"autodesk.unit.unit:kiloamperes", "KILOAMPERES"},
            {"autodesk.unit.unit:kilocalories", "KILOCALORIES"},
            {"autodesk.unit.unit:kilocaloriesPerSecond", "KILOCALORIES_PER_SECOND"},
            {"autodesk.unit.unit:kilogramForceMeters", "KILOGRAM_FORCE_METERS"},
            {"autodesk.unit.unit:kilogramForceMetersPerMeter", "KILOGRAM_FORCE_METERS_PER_METER"},
            {"autodesk.unit.unit:kilogramKelvins", "KILOGRAM_KELVINS"},
            {"autodesk.unit.unit:kilograms", "KILOGRAMS"},
            {"autodesk.unit.unit:kilogramsForce", "KILOGRAMS_FORCE"},
            {"autodesk.unit.unit:kilogramsForcePerMeter", "KILOGRAMS_FORCE_PER_METER"},
            {"autodesk.unit.unit:kilogramsForcePerSquareMeter", "KILOGRAMS_FORCE_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilogramsPerCubicMeter", "KILOGRAMS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:kilogramsPerHour", "KILOGRAMS_PER_HOUR"},
            {"autodesk.unit.unit:kilogramsPerKilogramKelvin", "KILOGRAMS_PER_KILOGRAM_KELVIN"},
            {"autodesk.unit.unit:kilogramsPerMeter", "KILOGRAMS_PER_METER"},
            {"autodesk.unit.unit:kilogramsPerMeterHour", "KILOGRAMS_PER_METER_HOUR"},
            {"autodesk.unit.unit:kilogramsPerMeterSecond", "KILOGRAMS_PER_METER_SECOND"},
            {"autodesk.unit.unit:kilogramsPerMinute", "KILOGRAMS_PER_MINUTE"},
            {"autodesk.unit.unit:kilogramsPerSecond", "KILOGRAMS_PER_SECOND"},
            {"autodesk.unit.unit:kilogramsPerSquareMeter", "KILOGRAMS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilojoules", "KILOJOULES"},
            {"autodesk.unit.unit:kilojoulesPerKelvin", "KILOJOULES_PER_KELVIN"},
            {"autodesk.unit.unit:kilometers", "KILOMETERS"},
            {"autodesk.unit.unit:kilometersPerHour", "KILOMETERS_PER_HOUR"},
            {"autodesk.unit.unit:kilometersPerSecond", "KILOMETERS_PER_SECOND"},
            {"autodesk.unit.unit:kilometersPerSecondSquared", "KILOMETERS_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:kilonewtonMeters", "KILONEWTON_METERS"},
            {"autodesk.unit.unit:kilonewtonMetersPerDegree", "KILONEWTON_METERS_PER_DEGREE"},
            {"autodesk.unit.unit:kilonewtonMetersPerDegreePerMeter", "KILONEWTON_METERS_PER_DEGREE_PER_METER"},
            {"autodesk.unit.unit:kilonewtonMetersPerMeter", "KILONEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:kilonewtons", "KILONEWTONS"},
            {"autodesk.unit.unit:kilonewtonsPerCubicMeter", "KILONEWTONS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:kilonewtonsPerMeter", "KILONEWTONS_PER_METER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareCentimeter", "KILONEWTONS_PER_SQUARE_CENTIMETER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareMeter", "KILONEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareMillimeter", "KILONEWTONS_PER_SQUARE_MILLIMETER"},
            {"autodesk.unit.unit:kilopascals", "KILOPASCALS"},
            {"autodesk.unit.unit:kilovoltAmperes", "KILOVOLT_AMPERES"},
            {"autodesk.unit.unit:kilovolts", "KILOVOLTS"},
            {"autodesk.unit.unit:kilowattHours", "KILOWATT_HOURS"},
            {"autodesk.unit.unit:kilowatts", "KILOWATTS"},
            {"autodesk.unit.unit:kipFeet", "KIP_FEET"},
            {"autodesk.unit.unit:kipFeetPerDegree", "KIP_FEET_PER_DEGREE"},
            {"autodesk.unit.unit:kipFeetPerDegreePerFoot", "KIP_FEET_PER_DEGREE_PER_FOOT"},
            {"autodesk.unit.unit:kipFeetPerFoot", "KIP_FEET_PER_FOOT"},
            {"autodesk.unit.unit:kips", "KIPS"},
            {"autodesk.unit.unit:kipsPerCubicFoot", "KIPS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:kipsPerCubicInch", "KIPS_PER_CUBIC_INCH"},
            {"autodesk.unit.unit:kipsPerFoot", "KIPS_PER_FOOT"},
            {"autodesk.unit.unit:kipsPerInch", "KIPS_PER_INCH"},
            {"autodesk.unit.unit:kipsPerSquareFoot", "KIPS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:kipsPerSquareInch", "KIPS_PER_SQUARE_INCH"},
            {"autodesk.unit.unit:liters", "LITERS"},
            {"autodesk.unit.unit:litersPerHour", "LITERS_PER_HOUR"},
            {"autodesk.unit.unit:litersPerMinute", "LITERS_PER_MINUTE"},
            {"autodesk.unit.unit:litersPerSecond", "LITERS_PER_SECOND"},
            {"autodesk.unit.unit:litersPerSecondCubicMeter", "LITERS_PER_SECOND_CUBIC_METER"},
            {"autodesk.unit.unit:litersPerSecondKilowatt", "LITERS_PER_SECOND_KILOWATT"},
            {"autodesk.unit.unit:litersPerSecondSquareMeter", "LITERS_PER_SECOND_SQUARE_METER"},
            {"autodesk.unit.unit:lumens", "LUMENS"},
            {"autodesk.unit.unit:lumensPerWatt", "LUMENS_PER_WATT"},
            {"autodesk.unit.unit:lux", "LUX"},
            {"autodesk.unit.unit:meganewtonMeters", "MEGANEWTON_METERS"},
            {"autodesk.unit.unit:meganewtonMetersPerMeter", "MEGANEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:meganewtons", "MEGANEWTONS"},
            {"autodesk.unit.unit:meganewtonsPerMeter", "MEGANEWTONS_PER_METER"},
            {"autodesk.unit.unit:meganewtonsPerSquareMeter", "MEGANEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:megapascals", "MEGAPASCALS"},
            {"autodesk.unit.unit:meters", "METERS"},
            {"autodesk.unit.unit:metersOfWaterColumn", "METERS_OF_WATER_COLUMN"},
            {"autodesk.unit.unit:metersOfWaterColumnPerMeter", "METERS_OF_WATER_COLUMN_PER_METER"},
            {"autodesk.unit.unit:metersPerKilonewton", "METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:metersPerSecond", "METERS_PER_SECOND"},
            {"autodesk.unit.unit:metersPerSecondSquared", "METERS_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:metersToTheFourthPower", "METERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:metersToTheSixthPower", "METERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:microinchesPerInchDegreeFahrenheit", "MICROINCHES_PER_INCH_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:micrometersPerMeterDegreeCelsius", "MICROMETERS_PER_METER_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:miles", "MILES"},
            {"autodesk.unit.unit:milesPerHour", "MILES_PER_HOUR"},
            {"autodesk.unit.unit:milesPerSecond", "MILES_PER_SECOND"},
            {"autodesk.unit.unit:milesPerSecondSquared", "MILES_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:milliamperes", "MILLIAMPERES"},
            {"autodesk.unit.unit:milligrams", "MILLIGRAMS"},
            {"autodesk.unit.unit:millimeters", "MILLIMETERS"},
            {"autodesk.unit.unit:millimetersOfMercury", "MILLIMETERS_OF_MERCURY"},
            {"autodesk.unit.unit:millimetersOfWaterColumn", "MILLIMETERS_OF_WATER_COLUMN"},
            {"autodesk.unit.unit:millimetersOfWaterColumnPerMeter", "MILLIMETERS_OF_WATER_COLUMN_PER_METER"},
            {"autodesk.unit.unit:millimetersToTheFourthPower", "MILLIMETERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:millimetersToTheSixthPower", "MILLIMETERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:milliseconds", "MILLISECONDS"},
            {"autodesk.unit.unit:millivolts", "MILLIVOLTS"},
            {"autodesk.unit.unit:minutes", "MINUTES"},
            {"autodesk.unit.unit:nanograms", "NANOGRAMS"},
            {"autodesk.unit.unit:nanogramsPerPascalSecondSquareMeter", "NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER"},
            {"autodesk.unit.unit:newtonMeters", "NEWTON_METERS"},
            {"autodesk.unit.unit:newtonMetersPerMeter", "NEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:newtonSecondsPerSquareMeter", "NEWTON_SECONDS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:newtons", "NEWTONS"},
            {"autodesk.unit.unit:newtonsPerMeter", "NEWTONS_PER_METER"},
            {"autodesk.unit.unit:newtonsPerSquareMeter", "NEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:newtonsPerSquareMillimeter", "NEWTONS_PER_SQUARE_MILLIMETER"},
            {"autodesk.unit.unit:ohmMeters", "OHM_METERS"},
            {"autodesk.unit.unit:ohms", "OHMS"},
            {"autodesk.unit.unit:pascalSeconds", "PASCAL_SECONDS"},
            {"autodesk.unit.unit:pascals", "PASCALS"},
            {"autodesk.unit.unit:pascalsPerMeter", "PASCALS_PER_METER"},
            {"autodesk.unit.unit:perMille", "PER_MILLE"},
            {"autodesk.unit.unit:percentage", "PERCENTAGE"},
            {"autodesk.unit.unit:pi", "MULTIPLES_OF_"},
            {"autodesk.unit.unit:poises", "POISES"},
            {"autodesk.unit.unit:poundForceFeet", "POUND_FORCE_FEET"},
            {"autodesk.unit.unit:poundForceFeetPerFoot", "POUND_FORCE_FEET_PER_FOOT"},
            {"autodesk.unit.unit:poundForceSecondsPerSquareFoot", "POUND_FORCE_SECONDS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:poundMassDegreesFahrenheit", "POUND_MASS_DEGREES_FAHRENHEIT"},
            {"autodesk.unit.unit:poundsForce", "POUNDS_FORCE"},
            {"autodesk.unit.unit:poundsForcePerCubicFoot", "POUNDS_FORCE_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:poundsForcePerFoot", "POUNDS_FORCE_PER_FOOT"},
            {"autodesk.unit.unit:poundsForcePerSquareFoot", "POUNDS_FORCE_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:poundsForcePerSquareInch", "POUNDS_FORCE_PER_SQUARE_INCH"},
            {"autodesk.unit.unit:poundsMass", "POUNDS_MASS"},
            {"autodesk.unit.unit:poundsMassPerCubicFoot", "POUNDS_MASS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:poundsMassPerCubicInch", "POUNDS_MASS_PER_CUBIC_INCH"},
            {"autodesk.unit.unit:poundsMassPerFoot", "POUNDS_MASS_PER_FOOT"},
            {"autodesk.unit.unit:poundsMassPerFootHour", "POUNDS_MASS_PER_FOOT_HOUR"},
            {"autodesk.unit.unit:poundsMassPerFootSecond", "POUNDS_MASS_PER_FOOT_SECOND"},
            {"autodesk.unit.unit:poundsMassPerHour", "POUNDS_MASS_PER_HOUR"},
            {"autodesk.unit.unit:poundsMassPerMinute", "POUNDS_MASS_PER_MINUTE"},
            {"autodesk.unit.unit:poundsMassPerPoundDegreeFahrenheit", "POUNDS_MASS_PER_POUND_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:poundsMassPerSecond", "POUNDS_MASS_PER_SECOND"},
            {"autodesk.unit.unit:poundsMassPerSquareFoot", "POUNDS_MASS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:radians", "RADIANS"},
            {"autodesk.unit.unit:radiansPerSecond", "RADIANS_PER_SECOND"},
            {"autodesk.unit.unit:rankine", "RANKINE"},
            {"autodesk.unit.unit:rankineInterval", "RANKINE_INTERVAL"},
            {"autodesk.unit.unit:ratioTo1", "RATIO_1"},
            {"autodesk.unit.unit:ratioTo10", "RATIO_10"},
            {"autodesk.unit.unit:ratioTo12", "RATIO_12"},
            {"autodesk.unit.unit:revolutionsPerMinute", "REVOLUTIONS_PER_MINUTE"},
            {"autodesk.unit.unit:revolutionsPerSecond", "REVOLUTIONS_PER_SECOND"},
            {"autodesk.unit.unit:riseDividedBy1000Millimeters", "RISE_1000_MILLIMETERS"},
            {"autodesk.unit.unit:riseDividedBy10Feet", "RISE_10_FEET"},
            {"autodesk.unit.unit:riseDividedBy120Inches", "RISE_120_INCHES"},
            {"autodesk.unit.unit:riseDividedBy12Inches", "RISE_12_INCHES"},
            {"autodesk.unit.unit:riseDividedBy1Foot", "RISE_1_FOOT"},
            {"autodesk.unit.unit:seconds", "SECONDS"},
            {"autodesk.unit.unit:squareCentimeters", "SQUARE_CENTIMETERS"},
            {"autodesk.unit.unit:squareCentimetersPerMeter", "SQUARE_CENTIMETERS_PER_METER"},
            {"autodesk.unit.unit:squareFeet", "SQUARE_FEET"},
            {"autodesk.unit.unit:squareFeetPer1000BritishThermalUnitsPerHour", "SQUARE_FEET_PER_THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:squareFeetPerFoot", "SQUARE_FEET_PER_FOOT"},
            {"autodesk.unit.unit:squareFeetPerKip", "SQUARE_FEET_PER_KIP"},
            {"autodesk.unit.unit:squareFeetPerSecond", "SQUARE_FEET_PER_SECOND"},
            {"autodesk.unit.unit:squareFeetPerTonOfRefrigeration", "SQUARE_FEET_PER_TON_OF_REFRIGERATION"},
            {"autodesk.unit.unit:squareHectometers", "SQUARE_HECTOMETERS"},
            {"autodesk.unit.unit:squareInches", "SQUARE_INCHES"},
            {"autodesk.unit.unit:squareInchesPerFoot", "SQUARE_INCHES_PER_FOOT"},
            {"autodesk.unit.unit:squareMeterKelvinsPerWatt", "SQUARE_METER_KELVINS_PER_WATT"},
            {"autodesk.unit.unit:squareMeters", "SQUARE_METERS"},
            {"autodesk.unit.unit:squareMetersPerKilonewton", "SQUARE_METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:squareMetersPerKilowatt", "SQUARE_METERS_PER_KILOWATT"},
            {"autodesk.unit.unit:squareMetersPerMeter", "SQUARE_METERS_PER_METER"},
            {"autodesk.unit.unit:squareMetersPerSecond", "SQUARE_METERS_PER_SECOND"},
            {"autodesk.unit.unit:squareMillimeters", "SQUARE_MILLIMETERS"},
            {"autodesk.unit.unit:squareMillimetersPerMeter", "SQUARE_MILLIMETERS_PER_METER"},
            {"autodesk.unit.unit:squareYards", "SQUARE_YARDS"},
            {"autodesk.unit.unit:standardGravity", "STANDARD_ACCELERATION_DUE_TO_GRAVITY"},
            {"autodesk.unit.unit:steradians", "STERADIANS"},
            {"autodesk.unit.unit:therms", "THERMS"},
            {"autodesk.unit.unit:thousandBritishThermalUnitsPerHour", "THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:tonneForceMeters", "TONNE_FORCE_METERS"},
            {"autodesk.unit.unit:tonneForceMetersPerMeter", "TONNE_FORCE_METERS_PER_METER"},
            {"autodesk.unit.unit:tonnes", "TONNES"},
            {"autodesk.unit.unit:tonnesForce", "TONNES_FORCE"},
            {"autodesk.unit.unit:tonnesForcePerMeter", "TONNES_FORCE_PER_METER"},
            {"autodesk.unit.unit:tonnesForcePerSquareMeter", "TONNES_FORCE_PER_SQUARE_METER"},
            {"autodesk.unit.unit:tonsOfRefrigeration", "TONS_OF_REFRIGERATION"},
            {"autodesk.unit.unit:turns", "TURNS"},
            {"autodesk.unit.unit:usGallons", "US_GALLONS"},
            {"autodesk.unit.unit:usGallonsPerHour", "US_GALLONS_PER_HOUR"},
            {"autodesk.unit.unit:usGallonsPerMinute", "US_GALLONS_PER_MINUTE"},
            {"autodesk.unit.unit:usSurveyFeet", "US_SURVEY_FEET"},
            {"autodesk.unit.unit:usTonnesForce", "US_TONNES_FORCE"},
            {"autodesk.unit.unit:usTonnesMass", "US_TONNES_MASS"},
            {"autodesk.unit.unit:voltAmperes", "VOLT_AMPERES"},
            {"autodesk.unit.unit:volts", "VOLTS"},
            {"autodesk.unit.unit:waterDensity4DegreesCelsius", "WATER_DENSITY_AT_4_DEGREES_CELSIUS"},
            {"autodesk.unit.unit:watts", "WATTS"},
            {"autodesk.unit.unit:wattsPerCubicFoot", "WATTS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:wattsPerCubicFootPerMinute", "WATTS_PER_CUBIC_FOOT_PER_MINUTE"},
            {"autodesk.unit.unit:wattsPerCubicMeter", "WATTS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:wattsPerCubicMeterPerSecond", "WATTS_PER_CUBIC_METER_PER_SECOND"},
            {"autodesk.unit.unit:wattsPerFoot", "WATTS_PER_FOOT"},
            {"autodesk.unit.unit:wattsPerMeter", "WATTS_PER_METER"},
            {"autodesk.unit.unit:wattsPerMeterKelvin", "WATTS_PER_METER_KELVIN"},
            {"autodesk.unit.unit:wattsPerSquareFoot", "WATTS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:wattsPerSquareMeter", "WATTS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:wattsPerSquareMeterKelvin", "WATTS_PER_SQUARE_METER_KELVIN"},
            {"autodesk.unit.unit:yards", "YARDS"},
            {"autodesk.unit.unit:feetFractionalInches", "FEET_AND_FRACTIONAL_INCHES"},
            {"autodesk.unit.unit:fractionalInches", "FRACTIONAL_INCHES"},
            {"autodesk.unit.unit:metersCentimeters", "METERS_AND_CENTIMETERS"},
            {"autodesk.unit.unit:degreesMinutes", "DEGREES_MINUTES_SECONDS"},
            {"autodesk.unit.unit:slopeDegrees", "SLOPE_DEGREES"},
            {"autodesk.unit.unit:stationingFeet", "FEET"},
            {"autodesk.unit.unit:stationingMeters", "METERS"},
            {"autodesk.unit.unit:stationingSurveyFeet", "US_SURVEY_FEET"},
            {"autodesk.unit.unit:ampereHours", "AMPERE_HOURS"},
            {"autodesk.unit.unit:ampereSeconds", "AMPERE_SECONDS"},
            {"autodesk.unit.unit:circularMils", "CIRCULAR_MILS"},
            {"autodesk.unit.unit:coulombs", "COULOMBS"},
            {"autodesk.unit.unit:dynes", "DYNES"},
            {"autodesk.unit.unit:ergs", "ERGS"},
            {"autodesk.unit.unit:farads", "FARADS"},
            {"autodesk.unit.unit:feetPerKipFoot", "FEET_PER_KIP_FOOT"},
            {"autodesk.unit.unit:gammas", "GAMMAS"},
            {"autodesk.unit.unit:gauss", "GAUSS"},
            {"autodesk.unit.unit:henries", "HENRIES"},
            {"autodesk.unit.unit:maxwells", "MAXWELLS"},
            {"autodesk.unit.unit:metersPerKilonewtonMeter", "METERS_PER_KILONEWTON_METER"},
            {"autodesk.unit.unit:mhos", "MHOS"},
            {"autodesk.unit.unit:microns", "MICRONS"},
            {"autodesk.unit.unit:mils", "MILS"},
            {"autodesk.unit.unit:nauticalMiles", "NAUTICAL_MILES"},
            {"autodesk.unit.unit:oersteds", "OERSTEDS"},
            {"autodesk.unit.unit:ouncesForce", "OUNCES_FORCE"},
            {"autodesk.unit.unit:ouncesMass", "OUNCES_MASS"},
            {"autodesk.unit.unit:siemens", "SIEMENS"},
            {"autodesk.unit.unit:slugs", "SLUGS"},
            {"autodesk.unit.unit:squareFeetPerKipFoot", "SQUARE_FEET_PER_KIP_FOOT"},
            {"autodesk.unit.unit:squareMetersPerKilonewtonMeter", "SQUARE_METERS_PER_KILONEWTON_METER"},
            {"autodesk.unit.unit:squareMils", "SQUARE_MILS"},
            {"autodesk.unit.unit:webers", "WEBERS"}
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (!doc.IsFamilyDocument)
            {
                MessageBox.Show(
                    "Необходимо открыть семейство в редакторе семейств",
                    "Экспорт CSV: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return Result.Cancelled;
            }
            Family family = doc.OwnerFamily;
            FamilySizeTableManager familySizeTableManager = FamilySizeTableManager.GetFamilySizeTableManager(doc, family.Id);

            IList<string> sizeTableNames = familySizeTableManager.GetAllSizeTableNames().ToList();
            if (sizeTableNames.Count == 0)
            {
                TaskDialog.Show("Экспорт CSV", "В семействе отсутствуют таблицы поиска");
                return Result.Cancelled;
            }

            // Выбор таблицы
            var selObjColl = SelectFromList("Выберите csv-таблицу для экспорта", sizeTableNames);
            if (selObjColl == null || selObjColl.Count == 0)
                return Result.Cancelled;

            Dictionary<string, string> tableDataDict = new Dictionary<string, string>();
            foreach (var obj in selObjColl)
            {
                string selectedSizeTableName = obj.ToString();
                if (string.IsNullOrEmpty(selectedSizeTableName))
                {
                    MessageBox.Show(
                        "Не удалось получить таблицу. Скинь разработчику",
                        "Экспорт CSV: Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    return Result.Cancelled;
                }

                FamilySizeTable selectedSizeTable = familySizeTableManager.GetSizeTable(selectedSizeTableName);
                int rows = selectedSizeTable.NumberOfRows;
                int columns = selectedSizeTable.NumberOfColumns;

                // Получение шапки таблицы
                StringBuilder returnedHeader = new StringBuilder(";");
                int versionNumber = int.Parse(doc.Application.VersionNumber);

                for (int i = 1; i < columns; i++)
                {
                    string resultForgeType = string.Empty;
                    string resultUnitType = string.Empty;

                    FamilySizeTableColumn columnHeader = selectedSizeTable.GetColumnHeader(i);
#if Revit2020 || Debug2020
                    if (versionNumber <= 2020)
                    {
                        resultForgeType = columnHeader.UnitType.ToString().Replace("UT_", "");
                        if (resultForgeType.Equals("Undefined"))
                            resultForgeType = "OTHER";
                        
                        resultUnitType = columnHeader.DisplayUnitType.ToString().Replace("DUT_", "");
                        if (resultUnitType.Equals("UNDEFINED"))
                            resultUnitType = string.Empty;
                    }
#else
                    if (versionNumber >= 2021)
                    {
                        try
                        {
                            string typeId = columnHeader.GetSpecTypeId().TypeId;
                            if (string.IsNullOrEmpty(typeId))
                            {
                                resultForgeType = "OTHER";
                                resultUnitType = string.Empty;
                            }
                            else
                            {
                                resultForgeType = GetValueFromDict_ByKeyStartWith(_forgeTypeIdDict, columnHeader.GetSpecTypeId().TypeId);
                                resultUnitType = GetValueFromDict_ByKeyStartWith(_unitTypeIdDict, columnHeader.GetUnitTypeId().TypeId);
                            }
                        }
                        catch
                        {
                            returnedHeader.AppendFormat("{0}##Undefined##UNDEFINED;", columnHeader.Name);
                        }
                    }
#endif
                    returnedHeader.AppendFormat(
                        "{0}##{1}##{2};",
                        columnHeader.Name,
                        resultForgeType,
                        resultUnitType);
                }


                returnedHeader.Length--; 
                returnedHeader.AppendLine();
                
                // Ручная замена неприемлемых форматов
                returnedHeader.Replace("##Undefined##UNDEFINED", "##OTHER##")
                              .Replace("##Angle##DEGREES_AND_MINUTES", "##Angle##DEGREES")
                              .Replace("DECIMAL_DEGREES", "DECIMAL DEGREES")
                              .Replace("DECIMAL_FEET", "GENERAL")
                              .Replace("Airflow", "AIR_FLOW");

                // Перебор строк таблицы
                StringBuilder returnedString = new StringBuilder(returnedHeader.ToString());

                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        returnedString.AppendFormat("{0};", selectedSizeTable.AsValueString(row, column).ToString());
                    }
                    returnedString.Length--; // Удалить последний символ ';'
                    returnedString.AppendLine();
                }

                tableDataDict[selectedSizeTableName] = returnedString.ToString();
            }

            // Запись файла
            string pathFolder = SelectFolder();
            if (string.IsNullOrEmpty(pathFolder))
            {
                MessageBox.Show(
                    "Не выбран путь для сохранения",
                    "Экспорт CSV: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                
                return Result.Cancelled;
            }

            foreach(KeyValuePair<string, string> kvp in tableDataDict)
            {
                string pathFile = Path.Combine(pathFolder, $"{kvp.Key}.csv");
                File.WriteAllText(pathFile, kvp.Value, Encoding.UTF8);
                ConvertEncoding(pathFile, Encoding.UTF8, Encoding.GetEncoding("windows-1251"));
            }

            TaskDialog.Show("Экспорт CSV", "Сохранение прошло успешно!");
            return Result.Succeeded;
        }

        private static ListBox.SelectedObjectCollection SelectFromList(string title, IList<string> options)
        {
            using (System.Windows.Forms.Form form = new System.Windows.Forms.Form())
            {
                form.Text = title;
                
                // Добавление элементов в окно
                ListBox listBox = new ListBox
                {
                    Dock = DockStyle.Fill,
                    SelectionMode = SelectionMode.MultiExtended,
                };
                listBox.Items.AddRange(options.ToArray());
                form.Controls.Add(listBox);
                
                // Обработчик события двойного нажатия на элемент списка
                listBox.DoubleClick += (sender, e) =>
                {
                    if (listBox.SelectedItem != null) // Убедимся, что что-то выбрано
                    {
                        form.DialogResult = DialogResult.OK;
                        form.Close();
                    }
                };

                // Добавление кнопки Выбрать
                Button button = new Button
                {
                    Text = "Выбрать",
                    Dock = DockStyle.Bottom
                };
                form.Controls.Add(button);

                // Обработчик события нажатия кнопки "Выбрать"
                button.Click += (sender, e) => 
                { 
                    form.DialogResult = DialogResult.OK; 
                    form.Close(); 
                };

                return form.ShowDialog() == DialogResult.OK ? listBox.SelectedItems : null;
            }
        }

        private static string SelectFolder()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                return folderBrowserDialog.ShowDialog() == DialogResult.OK ? folderBrowserDialog.SelectedPath : null;
            }
        }

        private static void ConvertEncoding(string filePath, Encoding sourceEncoding, Encoding targetEncoding)
        {
            string content = File.ReadAllText(filePath, sourceEncoding);
            File.WriteAllText(filePath, content, targetEncoding);
        }

        /// <summary>
        /// Получить Value из словаря по ключу словаря, только при условии "Начинается с" 
        /// </summary>
        /// <returns></returns>
        private static string GetValueFromDict_ByKeyStartWith(Dictionary<string, string> typeIdDict, string searchText)
        {
            string result = string.Empty;
            foreach (var kvp in typeIdDict)
            {
                if (searchText.StartsWith(kvp.Key))
                {
                    result = kvp.Value;
                    break;
                }
            }

            return result;
        }
    }
}
