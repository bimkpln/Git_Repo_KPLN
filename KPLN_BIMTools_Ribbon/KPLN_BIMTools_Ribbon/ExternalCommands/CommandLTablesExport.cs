using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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

        private readonly Dictionary<string, string> _forgeTypeIdDict = new Dictionary<string, string>
        {
            {"autodesk.spec.aec.electrical:apparentPower-1.0.0", "ELECTRICAL_APPARENT_POWER"},
            {"autodesk.spec.aec.electrical:cableTraySize-1.0.0", "CABLETRAY_SIZE"},
            {"autodesk.spec.aec.electrical:colorTemperature-1.0.0", "COLOR_TEMPERATURE"},
            {"autodesk.spec.aec.electrical:conduitSize-1.0.0", "CONDUIT_SIZE"},
            {"autodesk.spec.aec.electrical:costRateEnergy-1.0.0", "ELECTRICAL_COST_RATE_ENERGY"},
            {"autodesk.spec.aec.electrical:costRatePower-1.0.0", "ELECTRICAL_COST_RATE_POWER"},
            {"autodesk.spec.aec.electrical:current-1.0.0", "ELECTRICAL_CURRENT"},
            {"autodesk.spec.aec.electrical:demandFactor-1.0.0", "ELECTRICAL_DEMAND_FACTOR"},
            {"autodesk.spec.aec.electrical:efficacy-1.0.0", "ELECTRICAL_EFFICACY"},
            {"autodesk.spec.aec.electrical:frequency-1.0.0", "ELECTRICAL_FREQUENCY"},
            {"autodesk.spec.aec.electrical:illuminance-1.0.0", "ELECTRICAL_ILLUMINANCE"},
            {"autodesk.spec.aec.electrical:luminance-1.0.1", "ELECTRICAL_LUMINANCE"},
            {"autodesk.spec.aec.electrical:luminousFlux-1.0.0", "ELECTRICAL_LUMINOUS_FLUX"},
            {"autodesk.spec.aec.electrical:luminousIntensity-1.0.0", "ELECTRICAL_LUMINOUS_INTENSITY"},
            {"autodesk.spec.aec.electrical:potential-1.0.0", "ELECTRICAL_POTENTIAL"},
            {"autodesk.spec.aec.electrical:power-1.0.1", "ELECTRICAL_POWER"},
            {"autodesk.spec.aec.electrical:powerDensity-1.0.0", "ELECTRICAL_POWER_DENSITY"},
            {"autodesk.spec.aec.electrical:powerPerLength-1.0.0", "ELECTRICAL_POWER_PER_LENGTH"},
            {"autodesk.spec.aec.electrical:resistivity-1.0.0", "ELECTRICAL_RESISTIVITY"},
            {"autodesk.spec.aec.electrical:temperature-1.0.0", "ELECTRICAL_TEMPERATURE"},
            {"autodesk.spec.aec.electrical:temperatureDifference-1.0.0", "ELECTRICAL_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.electrical:wattage-1.0.0", "ELECTRICAL_WATTAGE"},
            {"autodesk.spec.aec.electrical:wireDiameter-1.0.0", "WIRE_SIZE"},
            {"autodesk.spec.aec.energy:energy-1.0.0", "HVAC_ENERGY"},
            {"autodesk.spec.aec.energy:heatTransferCoefficient-1.0.0", "HVAC_COEFFICIENT_OF_HEAT_TRANSFER"},
            {"autodesk.spec.aec.energy:isothermalMoistureCapacity-1.0.0", "HVAC_ISOTHERMAL_MOISTURE_CAPACITY"},
            {"autodesk.spec.aec.energy:permeability-1.0.0", "HVAC_PERMEABILITY"},
            {"autodesk.spec.aec.energy:specificHeat-1.0.0", "HVAC_SPECIFIC_HEAT"},
            {"autodesk.spec.aec.energy:specificHeatOfVaporization-1.0.0", "HVAC_SPECIFIC_HEAT_OF_VAPORIZATION"},
            {"autodesk.spec.aec.energy:thermalConductivity-1.0.0", "HVAC_THERMAL_CONDUCTIVITY"},
            {"autodesk.spec.aec.energy:thermalGradientCoefficientForMoistureCapacity-1.0.0", "HVAC_THERMAL_GRADIENT_COEFFICIENT_FOR_MOISTURE_CAPACITY"},
            {"autodesk.spec.aec.energy:thermalMass-1.0.0", "HVAC_THERMAL_MASS"},
            {"autodesk.spec.aec.energy:thermalResistance-1.0.0", "HVAC_THERMAL_RESISTANCE"},
            {"autodesk.spec.aec.hvac:airFlow-1.0.1", "HVAC_AIR_FLOW"},
            {"autodesk.spec.aec.hvac:airFlowDensity-1.0.1", "HVAC_AIRFLOW_DENSITY"},
            {"autodesk.spec.aec.hvac:airFlowDividedByCoolingLoad-1.0.0", "HVAC_AIRFLOW_DIVIDED_BY_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:airFlowDividedByVolume-1.0.1", "HVAC_AIRFLOW_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:angularSpeed-1.0.1", "HVAC_ANGULAR_SPEED"},
            {"autodesk.spec.aec.hvac:areaDividedByCoolingLoad-1.0.0", "HVAC_AREA_DIVIDED_BY_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:areaDividedByHeatingLoad-1.0.1", "HVAC_AREA_DIVIDED_BY_HEATING_LOAD"},
            {"autodesk.spec.aec.hvac:coolingLoad-1.0.0", "HVAC_COOLING_LOAD"},
            {"autodesk.spec.aec.hvac:coolingLoadDividedByArea-1.0.0", "HVAC_COOLING_LOAD_DIVIDED_BY_AREA"},
            {"autodesk.spec.aec.hvac:coolingLoadDividedByVolume-1.0.0", "HVAC_COOLING_LOAD_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:crossSection-1.0.0", "HVAC_CROSS_SECTION"},
            {"autodesk.spec.aec.hvac:density-1.0.0", "HVAC_DENSITY"},
            {"autodesk.spec.aec.hvac:diffusivity-1.0.0", "HVAC_DIFFUSIVITY"},
            {"autodesk.spec.aec.hvac:ductInsulationThickness-1.0.0", "HVAC_DUCT_INSULATION_THICKNESS"},
            {"autodesk.spec.aec.hvac:ductLiningThickness-1.0.0", "HVAC_DUCT_LINING_THICKNESS"},
            {"autodesk.spec.aec.hvac:ductSize-1.0.0", "HVAC_DUCT_SIZE"},
            {"autodesk.spec.aec.hvac:factor-1.0.0", "HVAC_FACTOR"},
            {"autodesk.spec.aec.hvac:flowPerPower-1.0.0", "HVAC_FLOW_PER_POWER"},
            {"autodesk.spec.aec.hvac:friction-1.0.2", "HVAC_FRICTION"},
            {"autodesk.spec.aec.hvac:heatGain-1.0.0", "HVAC_HEAT_GAIN"},
            {"autodesk.spec.aec.hvac:heatingLoad-1.0.0", "HVAC_HEATING_LOAD"},
            {"autodesk.spec.aec.hvac:heatingLoadDividedByArea-1.0.0", "HVAC_HEATING_LOAD_DIVIDED_BY_AREA"},
            {"autodesk.spec.aec.hvac:heatingLoadDividedByVolume-1.0.0", "HVAC_HEATING_LOAD_DIVIDED_BY_VOLUME"},
            {"autodesk.spec.aec.hvac:massPerTime-1.0.0", "HVAC_MASS_PER_TIME"},
            {"autodesk.spec.aec.hvac:power-1.0.0", "HVAC_POWER"},
            {"autodesk.spec.aec.hvac:powerDensity-1.0.0", "HVAC_POWER_DENSITY"},
            {"autodesk.spec.aec.hvac:powerPerFlow-1.0.0", "HVAC_POWER_PER_FLOW"},
            {"autodesk.spec.aec.hvac:pressure-1.0.2", "HVAC_PRESSURE"},
            {"autodesk.spec.aec.hvac:roughness-1.0.0", "HVAC_ROUGHNESS"},
            {"autodesk.spec.aec.hvac:slope-1.0.2", "HVAC_SLOPE"},
            {"autodesk.spec.aec.hvac:temperature-1.0.0", "HVAC_TEMPERATURE"},
            {"autodesk.spec.aec.hvac:temperatureDifference-1.0.0", "HVAC_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.hvac:velocity-1.0.1", "HVAC_VELOCITY"},
            {"autodesk.spec.aec.hvac:viscosity-1.0.1", "HVAC_VISCOSITY"},
            {"autodesk.spec.aec.infrastructure:stationing-1.0.1", "STATIONING"},
            {"autodesk.spec.aec.infrastructure:stationingInterval-1.0.0", "STATIONING_INTERVAL"},
            {"autodesk.spec.aec.piping:density-1.0.0", "PIPING_DENSITY"},
            {"autodesk.spec.aec.piping:flow-1.0.0", "PIPING_FLOW"},
            {"autodesk.spec.aec.piping:friction-1.0.2", "PIPING_FRICTION"},
            {"autodesk.spec.aec.piping:mass-1.0.1", "PIPE_MASS"},
            {"autodesk.spec.aec.piping:massPerTime-1.0.0", "PIPING_MASS_PER_TIME"},
            {"autodesk.spec.aec.piping:pipeDimension-1.0.0", "PIPE_DIMENSION"},
            {"autodesk.spec.aec.piping:pipeInsulationThickness-1.0.0", "PIPE_INSUlATION_THICKNESS"},
            {"autodesk.spec.aec.piping:pipeMassPerUnitLength-1.0.0", "PIPE_MASS_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.piping:pipeSize-1.0.0", "PIPE_SIZE"},
            {"autodesk.spec.aec.piping:pressure-1.0.2", "PIPING_PRESSURE"},
            {"autodesk.spec.aec.piping:roughness-1.0.0", "PIPING_ROUGHNESS"},
            {"autodesk.spec.aec.piping:slope-1.0.1", "PIPING_SLOPE"},
            {"autodesk.spec.aec.piping:temperature-1.0.0", "PIPING_TEMPERATURE"},
            {"autodesk.spec.aec.piping:temperatureDifference-1.0.0", "PIPING_TEMPERATURE_DIFFERENCE"},
            {"autodesk.spec.aec.piping:velocity-1.0.1", "PIPING_VELOCITY"},
            {"autodesk.spec.aec.piping:viscosity-1.0.1", "PIPING_VISCOSITY"},
            {"autodesk.spec.aec.piping:volume-1.0.0", "PIPING_VOLUME"},
            {"autodesk.spec.aec.structural:acceleration-1.0.0", "ACCELERATION"},
            {"autodesk.spec.aec.structural:areaForce-1.0.0", "AREA_FORCE"},
            {"autodesk.spec.aec.structural:areaForceScale-1.0.0", "AREA_FORCE_SCALE"},
            {"autodesk.spec.aec.structural:areaSpringCoefficient-1.0.0", "AREA_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:barDiameter-1.0.0", "BAR_DIAMETER"},
            {"autodesk.spec.aec.structural:crackWidth-1.0.0", "CRACK_WIDTH"},
            {"autodesk.spec.aec.structural:displacement-1.0.0", "DISPLACEMENT/DEFLECTION"},
            {"autodesk.spec.aec.structural:energy-1.0.0", "ENERGY"},
            {"autodesk.spec.aec.structural:force-1.0.1", "FORCE"},
            {"autodesk.spec.aec.structural:forceScale-1.0.0", "FORCE_SCALE"},
            {"autodesk.spec.aec.structural:frequency-1.0.0", "STRUCTURAL_FREQUENCY"},
            {"autodesk.spec.aec.structural:lineSpringCoefficient-1.0.0", "LINEAR_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:linearForce-1.0.0", "LINEAR_FORCE"},
            {"autodesk.spec.aec.structural:linearForceScale-1.0.0", "LINEAR_FORCE_SCALE"},
            {"autodesk.spec.aec.structural:linearMoment-1.0.0", "LINEAR_MOMENT"},
            {"autodesk.spec.aec.structural:linearMomentScale-1.0.0", "LINEAR_MOMENT_SCALE"},
            {"autodesk.spec.aec.structural:mass-1.0.0", "MASS"},
            {"autodesk.spec.aec.structural:massPerUnitArea-1.0.0", "MASS_PER_UNIT_AREA"},
            {"autodesk.spec.aec.structural:massPerUnitLength-1.0.0", "MASS_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.structural:moment-1.0.1", "MOMENT"},
            {"autodesk.spec.aec.structural:momentOfInertia-1.0.0", "MOMENT_OF_INERTIA"},
            {"autodesk.spec.aec.structural:momentScale-1.0.0", "MOMENT_SCALE"},
            {"autodesk.spec.aec.structural:period-1.0.0", "PERIOD"},
            {"autodesk.spec.aec.structural:pointSpringCoefficient-1.0.0", "POINT_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:pulsation-1.0.0", "PULSATION"},
            {"autodesk.spec.aec.structural:reinforcementArea-1.0.0", "REINFORCEMENT_AREA"},
            {"autodesk.spec.aec.structural:reinforcementAreaPerUnitLength-1.0.0", "REINFORCEMENT_AREA_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec.structural:reinforcementCover-1.0.0", "REINFORCEMENT_COVER"},
            {"autodesk.spec.aec.structural:reinforcementLength-1.0.0", "REINFORCEMENT_LENGTH"},
            {"autodesk.spec.aec.structural:reinforcementSpacing-1.0.0", "REINFORCEMENT_SPACING"},
            {"autodesk.spec.aec.structural:reinforcementVolume-1.0.0", "REINFORCEMENT_VOLUME"},
            {"autodesk.spec.aec.structural:rotation-1.0.0", "ROTATION"},
            {"autodesk.spec.aec.structural:rotationalLineSpringCoefficient-1.0.0", "ROTATIONAL_LINEAR_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:rotationalPointSpringCoefficient-1.0.0", "ROTATIONAL_POINT_SPRING_COEFFICIENT"},
            {"autodesk.spec.aec.structural:sectionArea-1.0.0", "SECTION_AREA"},
            {"autodesk.spec.aec.structural:sectionDimension-1.0.0", "SECTION_DIMENSION"},
            {"autodesk.spec.aec.structural:sectionModulus-1.0.0", "SECTION_MODULUS"},
            {"autodesk.spec.aec.structural:sectionProperty-1.0.0", "SECTION_PROPERTY"},
            {"autodesk.spec.aec.structural:stress-1.0.0", "STRESS"},
            {"autodesk.spec.aec.structural:surfaceAreaPerUnitLength-1.0.0", "SURFACE_AREA"},
            {"autodesk.spec.aec.structural:thermalExpansionCoefficient-1.0.0", "THERMAL_EXPANSION_COEFFICIENT"},
            {"autodesk.spec.aec.structural:unitWeight-1.0.0", "UNIT_WEIGHT"},
            {"autodesk.spec.aec.structural:velocity-1.0.1", "VELOCITY"},
            {"autodesk.spec.aec.structural:warpingConstant-1.0.0", "WARPING_CONSTANT"},
            {"autodesk.spec.aec.structural:weight-1.0.0", "WEIGHT"},
            {"autodesk.spec.aec.structural:weightPerUnitLength-1.0.0", "WEIGHT_PER_UNIT_LENGTH"},
            {"autodesk.spec.aec:angle-1.0.0", "ANGLE"},
            {"autodesk.spec.aec:area-1.0.0", "AREA"},
            {"autodesk.spec.aec:costPerArea-1.0.0", "COST_PER_AREA"},
            {"autodesk.spec.aec:decimalSheetLength-1.0.0", "PEN SIZE"},
            {"autodesk.spec.aec:distance-1.0.0", "DISTANCE"},
            {"autodesk.spec.aec:length-1.0.0", "LENGTH"},
            {"autodesk.spec.aec:massDensity-1.0.0", "MASS_DENSITY"},
            {"autodesk.spec.aec:number-1.0.1", "NUMBER"},
            {"autodesk.spec.aec:rotationAngle-1.0.1", "ROTATION_ANGLE"},
            {"autodesk.spec.aec:sheetLength-1.0.0", "PAPER LENGTH"},
            {"autodesk.spec.aec:siteAngle-1.0.0", "SITE_ANGLE"},
            {"autodesk.spec.aec:slope-1.0.1", "SLOPE"},
            {"autodesk.spec.aec:speed-1.0.1", "SPEED"},
            {"autodesk.spec.aec:time-1.0.0", "TIMEINTERVAL"},
            {"autodesk.spec.aec:volume-1.0.0", "VOLUME"},
            {"autodesk.spec.measurable:currency-1.0.0", "CURRENCY"}
        };

        private readonly Dictionary<string, string> _unitTypeIdDict = new Dictionary<string, string>
        {
            {"autodesk.unit.unit:1ToRatio-1.0.1", "RATIO1"},
            {"autodesk.unit.unit:acres-1.0.1", "ACRES"},
            {"autodesk.unit.unit:amperes-1.0.0", "AMPERES"},
            {"autodesk.unit.unit:atmospheres-1.0.1", "ATMOSPHERES"},
            {"autodesk.unit.unit:bars-1.0.1", "BARS"},
            {"autodesk.unit.unit:britishThermalUnits-1.0.1", "BRITISH_THERMAL_UNITS"},
            {"autodesk.unit.unit:britishThermalUnitsPerDegreeFahrenheit-1.0.1", "BRITISH_THERMAL_UNITS_PER_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHour-1.0.1", "BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourCubicFoot-1.0.1", "BRITISH_THERMAL_UNITS_PER_HOUR_CUBIC_FOOT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourFootDegreeFahrenheit-1.0.1", "BRITISH_THERMAL_UNITS_PER_HOUR_FOOT_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourSquareFoot-1.0.1", "BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT"},
            {"autodesk.unit.unit:britishThermalUnitsPerHourSquareFootDegreeFahrenheit-1.0.1", "BRITISH_THERMAL_UNITS_PER_HOUR_SQUARE_FOOT_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerPound-1.0.1", "BRITISH_THERMAL_UNITS_PER_POUND"},
            {"autodesk.unit.unit:britishThermalUnitsPerPoundDegreeFahrenheit-1.0.1", "BRITISH_THERMAL_UNITS_PER_POUND_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:britishThermalUnitsPerSecond-1.0.1", "BRITISH_THERMAL_UNITS_PER_SECOND"},
            {"autodesk.unit.unit:calories-1.0.1", "CALORIES"},
            {"autodesk.unit.unit:caloriesPerSecond-1.0.1", "CALORIES_PER_SECOND"},
            {"autodesk.unit.unit:candelas-1.0.0", "CANDELAS"},
            {"autodesk.unit.unit:candelasPerSquareFoot-1.0.0", "CANDELAS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:candelasPerSquareMeter-1.0.1", "CANDELAS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:celsius-1.0.1", "CELSIUS"},
            {"autodesk.unit.unit:celsiusInterval-1.0.1", "CELSIUS_INTERVAL"},
            {"autodesk.unit.unit:centimeters-1.0.1", "CENTIMETERS"},
            {"autodesk.unit.unit:centimetersPerMinute-1.0.1", "CENTIMETERS_PER_MINUTE"},
            {"autodesk.unit.unit:centimetersToTheFourthPower-1.0.1", "CENTIMETERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:centimetersToTheSixthPower-1.0.1", "CENTIMETERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:centipoises-1.0.1", "CENTIPOISES"},
            {"autodesk.unit.unit:cubicCentimeters-1.0.1", "CUBIC_CENTIMETERS"},
            {"autodesk.unit.unit:cubicFeet-1.0.1", "CUBIC_FEET"},
            {"autodesk.unit.unit:cubicFeetPerHour-1.0.0", "CUBIC_FEET_PER_HOUR"},
            {"autodesk.unit.unit:cubicFeetPerKip-1.0.1", "CUBIC_FEET_PER_KIP"},
            {"autodesk.unit.unit:cubicFeetPerMinute-1.0.1", "CUBIC_FEET_PER_MINUTE"},
            {"autodesk.unit.unit:cubicFeetPerMinuteCubicFoot-1.0.1", "CUBIC_FEET_PER_MINUTE_CUBIC_FOOT"},
            {"autodesk.unit.unit:cubicFeetPerMinutePerBritishThermalUnitPerHour-1.0.0", "CUBIC_FEET_PER_MINUTE_PER_BRITISH_THERMAL_UNIT_PER_HOUR"},
            {"autodesk.unit.unit:cubicFeetPerMinuteSquareFoot-1.0.1", "CUBIC_FEET_PER_MINUTE_SQUARE_FOOT"},
            {"autodesk.unit.unit:cubicFeetPerMinuteTonOfRefrigeration-1.0.1", "CUBIC_FEET_PER_MINUTE_TON_OF_REFRIGERATION"},
            {"autodesk.unit.unit:cubicFeetPerPoundMass-1.0.0", "CUBIC_FEET_PER_POUND_MASS"},
            {"autodesk.unit.unit:cubicInches-1.0.1", "CUBIC_INCHES"},
            {"autodesk.unit.unit:cubicMeters-1.0.1", "CUBIC_METERS"},
            {"autodesk.unit.unit:cubicMetersPerHour-1.0.1", "CUBIC_METERS_PER_HOUR"},
            {"autodesk.unit.unit:cubicMetersPerHourCubicMeter-1.0.0", "CUBIC_METERS_PER_HOUR_CUBIC_METER"},
            {"autodesk.unit.unit:cubicMetersPerHourSquareMeter-1.0.0", "CUBIC_METERS_PER_HOUR_SQUARE_METER"},
            {"autodesk.unit.unit:cubicMetersPerKilogram-1.0.0", "CUBIC_METERS_PER_KILOGRAM"},
            {"autodesk.unit.unit:cubicMetersPerKilonewton-1.0.1", "CUBIC_METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:cubicMetersPerSecond-1.0.1", "CUBIC_METERS_PER_SECOND"},
            {"autodesk.unit.unit:cubicMetersPerWattSecond-1.0.0", "CUBIC_METERS_PER_WATT_SECOND"},
            {"autodesk.unit.unit:cubicMillimeters-1.0.1", "CUBIC_MILLIMETERS"},
            {"autodesk.unit.unit:cubicYards-1.0.1", "CUBIC_YARDS"},
            {"autodesk.unit.unit:currency-1.0.0", "CURRENCY"},
            {"autodesk.unit.unit:currencyPerBritishThermalUnit-1.0.1", "COST_PER_BRITISH_THERMAL_UNIT"},
            {"autodesk.unit.unit:currencyPerBritishThermalUnitPerHour-1.0.1", "COST_PER_BRITISH_THERMAL_UNIT_PER_HOUR"},
            {"autodesk.unit.unit:currencyPerSquareFoot-1.0.1", "COST_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:currencyPerSquareMeter-1.0.1", "COST_PER_SQUARE_METER"},
            {"autodesk.unit.unit:currencyPerWatt-1.0.1", "COST_PER_WATT"},
            {"autodesk.unit.unit:currencyPerWattHour-1.0.1", "COST_PER_WATT_HOUR"},
            {"autodesk.unit.unit:cyclesPerSecond-1.0.1", "CYCLES_PER_SECOND"},
            {"autodesk.unit.unit:decimeters-1.0.1", "DECIMETERS"},
            {"autodesk.unit.unit:degrees-1.0.1", "DEGREES"},
            {"autodesk.unit.unit:dekanewtonMeters-1.0.1", "DEKANEWTON_METERS"},
            {"autodesk.unit.unit:dekanewtonMetersPerMeter-1.0.1", "DEKANEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:dekanewtons-1.0.1", "DEKANEWTONS"},
            {"autodesk.unit.unit:dekanewtonsPerMeter-1.0.1", "DEKANEWTONS_PER_METER"},
            {"autodesk.unit.unit:dekanewtonsPerSquareMeter-1.0.1", "DEKANEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:fahrenheit-1.0.1", "FAHRENHEIT"},
            {"autodesk.unit.unit:fahrenheitInterval-1.0.1", "FAHRENHEIT_INTERVAL"},
            {"autodesk.unit.unit:feet-1.0.1", "FEET"},
            {"autodesk.unit.unit:feetOfWater39.2DegreesFahrenheit-1.0.1", "FEET_OF_WATER"},
            {"autodesk.unit.unit:feetOfWater39.2DegreesFahrenheitPer100Feet-1.0.1", "FEET_OF_WATER_PER_100FT"},
            {"autodesk.unit.unit:feetPerKip-1.0.0", "FEET_PER_KIP"},
            {"autodesk.unit.unit:feetPerMinute-1.0.1", "FEET_PER_MINUTE"},
            {"autodesk.unit.unit:feetPerSecond-1.0.1", "FEET_PER_SECOND"},
            {"autodesk.unit.unit:feetPerSecondSquared-1.0.1", "FEET_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:feetToTheFourthPower-1.0.1", "FEET_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:feetToTheSixthPower-1.0.1", "FEET_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:fixed-1.0.1", "FIXED"},
            {"autodesk.unit.unit:footcandles-1.0.1", "FOOTCANDLES"},
            {"autodesk.unit.unit:footlamberts-1.0.1", "FOOTLAMBERTS"},
            {"autodesk.unit.unit:general-1.0.1", "GENERAL"},
            {"autodesk.unit.unit:gradians-1.0.1", "GRADIANS"},
            {"autodesk.unit.unit:grains-1.0.1", "GRAINS"},
            {"autodesk.unit.unit:grainsPerHourSquareFootInchMercury-1.0.1", "GRAINS_PER_HOUR_SQUARE_FOOT_INCH_MERCURY"},
            {"autodesk.unit.unit:grams-1.0.1", "GRAMS"},
            {"autodesk.unit.unit:hectares-1.0.1", "HECTARES"},
            {"autodesk.unit.unit:hectometers-1.0.1", "HECTOMETERS"},
            {"autodesk.unit.unit:hertz-1.0.1", "HERTZ"},
            {"autodesk.unit.unit:horsepower-1.0.1", "HORSEPOWER"},
            {"autodesk.unit.unit:hourSquareFootDegreesFahrenheitPerBritishThermalUnit-1.0.1", "HOUR_SQUARE_FOOT_DEGREES_FAHRENHEIT_PER_BRITISH_THERMAL_UNIT"},
            {"autodesk.unit.unit:hours-1.0.1", "HOURS"},
            {"autodesk.unit.unit:inches-1.0.1", "INCHES"},
            {"autodesk.unit.unit:inchesOfMercury32DegreesFahrenheit-1.0.1", "INCHES_OF_MERCURY"},
            {"autodesk.unit.unit:inchesOfWater60DegreesFahrenheit-1.0.1", "INCHES_OF_WATER"},
            {"autodesk.unit.unit:inchesOfWater60DegreesFahrenheitPer100Feet-1.0.1", "INCHES_OF_WATER_PER_100FT"},
            {"autodesk.unit.unit:inchesPerSecond-1.0.1", "INCHES_PER_SECOND"},
            {"autodesk.unit.unit:inchesPerSecondSquared-1.0.1", "INCHES_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:inchesToTheFourthPower-1.0.1", "INCHES_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:inchesToTheSixthPower-1.0.1", "INCHES_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:inverseDegreesCelsius-1.0.1", "INVERSE_DEGREES_CELSIUS"},
            {"autodesk.unit.unit:inverseDegreesFahrenheit-1.0.1", "INVERSE_DEGREES_FAHRENHEIT"},
            {"autodesk.unit.unit:inverseKilonewtons-1.0.0", "INVERSE_KILONEWTONS"},
            {"autodesk.unit.unit:inverseKips-1.0.0", "INVERSE_KIPS"},
            {"autodesk.unit.unit:joules-1.0.1", "JOULES"},
            {"autodesk.unit.unit:joulesPerGram-1.0.1", "JOULES_PER_GRAM"},
            {"autodesk.unit.unit:joulesPerGramDegreeCelsius-1.0.1", "JOULES_PER_GRAM_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:joulesPerKelvin-1.0.1", "JOULES_PER_KELVIN"},
            {"autodesk.unit.unit:joulesPerKilogram-1.0.1", "JOULES_PER_KILOGRAM"},
            {"autodesk.unit.unit:joulesPerKilogramDegreeCelsius-1.0.1", "JOULES_PER_KILOGRAM_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:kelvin-1.0.0", "KELVIN"},
            {"autodesk.unit.unit:kelvinInterval-1.0.0", "KELVIN_INTERVAL"},
            {"autodesk.unit.unit:kiloamperes-1.0.1", "KILOAMPERES"},
            {"autodesk.unit.unit:kilocalories-1.0.1", "KILOCALORIES"},
            {"autodesk.unit.unit:kilocaloriesPerSecond-1.0.1", "KILOCALORIES_PER_SECOND"},
            {"autodesk.unit.unit:kilogramForceMeters-1.0.1", "KILOGRAM_FORCE_METERS"},
            {"autodesk.unit.unit:kilogramForceMetersPerMeter-1.0.1", "KILOGRAM_FORCE_METERS_PER_METER"},
            {"autodesk.unit.unit:kilogramKelvins-1.0.0", "KILOGRAM_KELVINS"},
            {"autodesk.unit.unit:kilograms-1.0.0", "KILOGRAMS"},
            {"autodesk.unit.unit:kilogramsForce-1.0.1", "KILOGRAMS_FORCE"},
            {"autodesk.unit.unit:kilogramsForcePerMeter-1.0.1", "KILOGRAMS_FORCE_PER_METER"},
            {"autodesk.unit.unit:kilogramsForcePerSquareMeter-1.0.1", "KILOGRAMS_FORCE_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilogramsPerCubicMeter-1.0.1", "KILOGRAMS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:kilogramsPerHour-1.0.0", "KILOGRAMS_PER_HOUR"},
            {"autodesk.unit.unit:kilogramsPerKilogramKelvin-1.0.0", "KILOGRAMS_PER_KILOGRAM_KELVIN"},
            {"autodesk.unit.unit:kilogramsPerMeter-1.0.1", "KILOGRAMS_PER_METER"},
            {"autodesk.unit.unit:kilogramsPerMeterHour-1.0.0", "KILOGRAMS_PER_METER_HOUR"},
            {"autodesk.unit.unit:kilogramsPerMeterSecond-1.0.0", "KILOGRAMS_PER_METER_SECOND"},
            {"autodesk.unit.unit:kilogramsPerMinute-1.0.0", "KILOGRAMS_PER_MINUTE"},
            {"autodesk.unit.unit:kilogramsPerSecond-1.0.0", "KILOGRAMS_PER_SECOND"},
            {"autodesk.unit.unit:kilogramsPerSquareMeter-1.0.1", "KILOGRAMS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilojoules-1.0.1", "KILOJOULES"},
            {"autodesk.unit.unit:kilojoulesPerKelvin-1.0.1", "KILOJOULES_PER_KELVIN"},
            {"autodesk.unit.unit:kilometers-1.0.1", "KILOMETERS"},
            {"autodesk.unit.unit:kilometersPerHour-1.0.1", "KILOMETERS_PER_HOUR"},
            {"autodesk.unit.unit:kilometersPerSecond-1.0.1", "KILOMETERS_PER_SECOND"},
            {"autodesk.unit.unit:kilometersPerSecondSquared-1.0.1", "KILOMETERS_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:kilonewtonMeters-1.0.1", "KILONEWTON_METERS"},
            {"autodesk.unit.unit:kilonewtonMetersPerDegree-1.0.1", "KILONEWTON_METERS_PER_DEGREE"},
            {"autodesk.unit.unit:kilonewtonMetersPerDegreePerMeter-1.0.1", "KILONEWTON_METERS_PER_DEGREE_PER_METER"},
            {"autodesk.unit.unit:kilonewtonMetersPerMeter-1.0.1", "KILONEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:kilonewtons-1.0.1", "KILONEWTONS"},
            {"autodesk.unit.unit:kilonewtonsPerCubicMeter-1.0.1", "KILONEWTONS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:kilonewtonsPerMeter-1.0.1", "KILONEWTONS_PER_METER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareCentimeter-1.0.1", "KILONEWTONS_PER_SQUARE_CENTIMETER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareMeter-1.0.1", "KILONEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:kilonewtonsPerSquareMillimeter-1.0.1", "KILONEWTONS_PER_SQUARE_MILLIMETER"},
            {"autodesk.unit.unit:kilopascals-1.0.1", "KILOPASCALS"},
            {"autodesk.unit.unit:kilovoltAmperes-1.0.1", "KILOVOLT_AMPERES"},
            {"autodesk.unit.unit:kilovolts-1.0.1", "KILOVOLTS"},
            {"autodesk.unit.unit:kilowattHours-1.0.1", "KILOWATT_HOURS"},
            {"autodesk.unit.unit:kilowatts-1.0.1", "KILOWATTS"},
            {"autodesk.unit.unit:kipFeet-1.0.1", "KIP_FEET"},
            {"autodesk.unit.unit:kipFeetPerDegree-1.0.1", "KIP_FEET_PER_DEGREE"},
            {"autodesk.unit.unit:kipFeetPerDegreePerFoot-1.0.1", "KIP_FEET_PER_DEGREE_PER_FOOT"},
            {"autodesk.unit.unit:kipFeetPerFoot-1.0.1", "KIP_FEET_PER_FOOT"},
            {"autodesk.unit.unit:kips-1.0.1", "KIPS"},
            {"autodesk.unit.unit:kipsPerCubicFoot-1.0.1", "KIPS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:kipsPerCubicInch-1.0.1", "KIPS_PER_CUBIC_INCH"},
            {"autodesk.unit.unit:kipsPerFoot-1.0.1", "KIPS_PER_FOOT"},
            {"autodesk.unit.unit:kipsPerInch-1.0.1", "KIPS_PER_INCH"},
            {"autodesk.unit.unit:kipsPerSquareFoot-1.0.1", "KIPS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:kipsPerSquareInch-1.0.1", "KIPS_PER_SQUARE_INCH"},
            {"autodesk.unit.unit:liters-1.0.1", "LITERS"},
            {"autodesk.unit.unit:litersPerHour-1.0.1", "LITERS_PER_HOUR"},
            {"autodesk.unit.unit:litersPerMinute-1.0.1", "LITERS_PER_MINUTE"},
            {"autodesk.unit.unit:litersPerSecond-1.0.1", "LITERS_PER_SECOND"},
            {"autodesk.unit.unit:litersPerSecondCubicMeter-1.0.1", "LITERS_PER_SECOND_CUBIC_METER"},
            {"autodesk.unit.unit:litersPerSecondKilowatt-1.0.1", "LITERS_PER_SECOND_KILOWATT"},
            {"autodesk.unit.unit:litersPerSecondSquareMeter-1.0.1", "LITERS_PER_SECOND_SQUARE_METER"},
            {"autodesk.unit.unit:lumens-1.0.1", "LUMENS"},
            {"autodesk.unit.unit:lumensPerWatt-1.0.1", "LUMENS_PER_WATT"},
            {"autodesk.unit.unit:lux-1.0.1", "LUX"},
            {"autodesk.unit.unit:meganewtonMeters-1.0.1", "MEGANEWTON_METERS"},
            {"autodesk.unit.unit:meganewtonMetersPerMeter-1.0.1", "MEGANEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:meganewtons-1.0.1", "MEGANEWTONS"},
            {"autodesk.unit.unit:meganewtonsPerMeter-1.0.1", "MEGANEWTONS_PER_METER"},
            {"autodesk.unit.unit:meganewtonsPerSquareMeter-1.0.1", "MEGANEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:megapascals-1.0.1", "MEGAPASCALS"},
            {"autodesk.unit.unit:meters-1.0.0", "METERS"},
            {"autodesk.unit.unit:metersOfWaterColumn-1.0.0", "METERS_OF_WATER_COLUMN"},
            {"autodesk.unit.unit:metersOfWaterColumnPerMeter-1.0.0", "METERS_OF_WATER_COLUMN_PER_METER"},
            {"autodesk.unit.unit:metersPerKilonewton-1.0.0", "METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:metersPerSecond-1.0.1", "METERS_PER_SECOND"},
            {"autodesk.unit.unit:metersPerSecondSquared-1.0.1", "METERS_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:metersToTheFourthPower-1.0.1", "METERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:metersToTheSixthPower-1.0.1", "METERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:microinchesPerInchDegreeFahrenheit-1.0.1", "MICROINCHES_PER_INCH_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:micrometersPerMeterDegreeCelsius-1.0.1", "MICROMETERS_PER_METER_DEGREE_CELSIUS"},
            {"autodesk.unit.unit:miles-1.0.1", "MILES"},
            {"autodesk.unit.unit:milesPerHour-1.0.1", "MILES_PER_HOUR"},
            {"autodesk.unit.unit:milesPerSecond-1.0.1", "MILES_PER_SECOND"},
            {"autodesk.unit.unit:milesPerSecondSquared-1.0.1", "MILES_PER_SECOND_SQUARED"},
            {"autodesk.unit.unit:milliamperes-1.0.1", "MILLIAMPERES"},
            {"autodesk.unit.unit:milligrams-1.0.1", "MILLIGRAMS"},
            {"autodesk.unit.unit:millimeters-1.0.1", "MILLIMETERS"},
            {"autodesk.unit.unit:millimetersOfMercury-1.0.1", "MILLIMETERS_OF_MERCURY"},
            {"autodesk.unit.unit:millimetersOfWaterColumn-1.0.0", "MILLIMETERS_OF_WATER_COLUMN"},
            {"autodesk.unit.unit:millimetersOfWaterColumnPerMeter-1.0.0", "MILLIMETERS_OF_WATER_COLUMN_PER_METER"},
            {"autodesk.unit.unit:millimetersToTheFourthPower-1.0.1", "MILLIMETERS_TO_THE_FOURTH_POWER"},
            {"autodesk.unit.unit:millimetersToTheSixthPower-1.0.1", "MILLIMETERS_TO_THE_SIXTH_POWER"},
            {"autodesk.unit.unit:milliseconds-1.0.1", "MILLISECONDS"},
            {"autodesk.unit.unit:millivolts-1.0.1", "MILLIVOLTS"},
            {"autodesk.unit.unit:minutes-1.0.1", "MINUTES"},
            {"autodesk.unit.unit:nanograms-1.0.1", "NANOGRAMS"},
            {"autodesk.unit.unit:nanogramsPerPascalSecondSquareMeter-1.0.1", "NANOGRAMS_PER_PASCAL_SECOND_SQUARE_METER"},
            {"autodesk.unit.unit:newtonMeters-1.0.1", "NEWTON_METERS"},
            {"autodesk.unit.unit:newtonMetersPerMeter-1.0.1", "NEWTON_METERS_PER_METER"},
            {"autodesk.unit.unit:newtonSecondsPerSquareMeter-1.0.0", "NEWTON_SECONDS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:newtons-1.0.1", "NEWTONS"},
            {"autodesk.unit.unit:newtonsPerMeter-1.0.1", "NEWTONS_PER_METER"},
            {"autodesk.unit.unit:newtonsPerSquareMeter-1.0.1", "NEWTONS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:newtonsPerSquareMillimeter-1.0.1", "NEWTONS_PER_SQUARE_MILLIMETER"},
            {"autodesk.unit.unit:ohmMeters-1.0.1", "OHM_METERS"},
            {"autodesk.unit.unit:ohms-1.0.1", "OHMS"},
            {"autodesk.unit.unit:pascalSeconds-1.0.1", "PASCAL_SECONDS"},
            {"autodesk.unit.unit:pascals-1.0.1", "PASCALS"},
            {"autodesk.unit.unit:pascalsPerMeter-1.0.1", "PASCALS_PER_METER"},
            {"autodesk.unit.unit:perMille-1.0.1", "PER_MILLE"},
            {"autodesk.unit.unit:percentage-1.0.1", "PERCENTAGE"},
            {"autodesk.unit.unit:pi-1.0.0", "MULTIPLES_OF_"},
            {"autodesk.unit.unit:poises-1.0.1", "POISES"},
            {"autodesk.unit.unit:poundForceFeet-1.0.1", "POUND_FORCE_FEET"},
            {"autodesk.unit.unit:poundForceFeetPerFoot-1.0.1", "POUND_FORCE_FEET_PER_FOOT"},
            {"autodesk.unit.unit:poundForceSecondsPerSquareFoot-1.0.0", "POUND_FORCE_SECONDS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:poundMassDegreesFahrenheit-1.0.0", "POUND_MASS_DEGREES_FAHRENHEIT"},
            {"autodesk.unit.unit:poundsForce-1.0.1", "POUNDS_FORCE"},
            {"autodesk.unit.unit:poundsForcePerCubicFoot-1.0.1", "POUNDS_FORCE_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:poundsForcePerFoot-1.0.1", "POUNDS_FORCE_PER_FOOT"},
            {"autodesk.unit.unit:poundsForcePerSquareFoot-1.0.1", "POUNDS_FORCE_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:poundsForcePerSquareInch-1.0.1", "POUNDS_FORCE_PER_SQUARE_INCH"},
            {"autodesk.unit.unit:poundsMass-1.0.1", "POUNDS_MASS"},
            {"autodesk.unit.unit:poundsMassPerCubicFoot-1.0.1", "POUNDS_MASS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:poundsMassPerCubicInch-1.0.1", "POUNDS_MASS_PER_CUBIC_INCH"},
            {"autodesk.unit.unit:poundsMassPerFoot-1.0.1", "POUNDS_MASS_PER_FOOT"},
            {"autodesk.unit.unit:poundsMassPerFootHour-1.0.1", "POUNDS_MASS_PER_FOOT_HOUR"},
            {"autodesk.unit.unit:poundsMassPerFootSecond-1.0.1", "POUNDS_MASS_PER_FOOT_SECOND"},
            {"autodesk.unit.unit:poundsMassPerHour-1.0.0", "POUNDS_MASS_PER_HOUR"},
            {"autodesk.unit.unit:poundsMassPerMinute-1.0.0", "POUNDS_MASS_PER_MINUTE"},
            {"autodesk.unit.unit:poundsMassPerPoundDegreeFahrenheit-1.0.0", "POUNDS_MASS_PER_POUND_DEGREE_FAHRENHEIT"},
            {"autodesk.unit.unit:poundsMassPerSecond-1.0.0", "POUNDS_MASS_PER_SECOND"},
            {"autodesk.unit.unit:poundsMassPerSquareFoot-1.0.1", "POUNDS_MASS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:radians-1.0.0", "RADIANS"},
            {"autodesk.unit.unit:radiansPerSecond-1.0.1", "RADIANS_PER_SECOND"},
            {"autodesk.unit.unit:rankine-1.0.1", "RANKINE"},
            {"autodesk.unit.unit:rankineInterval-1.0.1", "RANKINE_INTERVAL"},
            {"autodesk.unit.unit:ratioTo1-1.0.0", "RATIO_1"},
            {"autodesk.unit.unit:ratioTo10-1.0.1", "RATIO_10"},
            {"autodesk.unit.unit:ratioTo12-1.0.1", "RATIO_12"},
            {"autodesk.unit.unit:revolutionsPerMinute-1.0.0", "REVOLUTIONS_PER_MINUTE"},
            {"autodesk.unit.unit:revolutionsPerSecond-1.0.0", "REVOLUTIONS_PER_SECOND"},
            {"autodesk.unit.unit:riseDividedBy1000Millimeters-1.0.1", "RISE_1000_MILLIMETERS"},
            {"autodesk.unit.unit:riseDividedBy10Feet-1.0.1", "RISE_10_FEET"},
            {"autodesk.unit.unit:riseDividedBy120Inches-1.0.1", "RISE_120_INCHES"},
            {"autodesk.unit.unit:riseDividedBy12Inches-1.0.1", "RISE_12_INCHES"},
            {"autodesk.unit.unit:riseDividedBy1Foot-1.0.1", "RISE_1_FOOT"},
            {"autodesk.unit.unit:seconds-1.0.0", "SECONDS"},
            {"autodesk.unit.unit:squareCentimeters-1.0.1", "SQUARE_CENTIMETERS"},
            {"autodesk.unit.unit:squareCentimetersPerMeter-1.0.1", "SQUARE_CENTIMETERS_PER_METER"},
            {"autodesk.unit.unit:squareFeet-1.0.1", "SQUARE_FEET"},
            {"autodesk.unit.unit:squareFeetPer1000BritishThermalUnitsPerHour-1.0.1", "SQUARE_FEET_PER_THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:squareFeetPerFoot-1.0.1", "SQUARE_FEET_PER_FOOT"},
            {"autodesk.unit.unit:squareFeetPerKip-1.0.1", "SQUARE_FEET_PER_KIP"},
            {"autodesk.unit.unit:squareFeetPerSecond-1.0.0", "SQUARE_FEET_PER_SECOND"},
            {"autodesk.unit.unit:squareFeetPerTonOfRefrigeration-1.0.1", "SQUARE_FEET_PER_TON_OF_REFRIGERATION"},
            {"autodesk.unit.unit:squareHectometers-1.0.1", "SQUARE_HECTOMETERS"},
            {"autodesk.unit.unit:squareInches-1.0.1", "SQUARE_INCHES"},
            {"autodesk.unit.unit:squareInchesPerFoot-1.0.1", "SQUARE_INCHES_PER_FOOT"},
            {"autodesk.unit.unit:squareMeterKelvinsPerWatt-1.0.1", "SQUARE_METER_KELVINS_PER_WATT"},
            {"autodesk.unit.unit:squareMeters-1.0.1", "SQUARE_METERS"},
            {"autodesk.unit.unit:squareMetersPerKilonewton-1.0.1", "SQUARE_METERS_PER_KILONEWTON"},
            {"autodesk.unit.unit:squareMetersPerKilowatt-1.0.1", "SQUARE_METERS_PER_KILOWATT"},
            {"autodesk.unit.unit:squareMetersPerMeter-1.0.1", "SQUARE_METERS_PER_METER"},
            {"autodesk.unit.unit:squareMetersPerSecond-1.0.0", "SQUARE_METERS_PER_SECOND"},
            {"autodesk.unit.unit:squareMillimeters-1.0.1", "SQUARE_MILLIMETERS"},
            {"autodesk.unit.unit:squareMillimetersPerMeter-1.0.1", "SQUARE_MILLIMETERS_PER_METER"},
            {"autodesk.unit.unit:squareYards-1.0.1", "SQUARE_YARDS"},
            {"autodesk.unit.unit:standardGravity-1.0.1", "STANDARD_ACCELERATION_DUE_TO_GRAVITY"},
            {"autodesk.unit.unit:steradians-1.0.0", "STERADIANS"},
            {"autodesk.unit.unit:therms-1.0.1", "THERMS"},
            {"autodesk.unit.unit:thousandBritishThermalUnitsPerHour-1.0.0", "THOUSAND_BRITISH_THERMAL_UNITS_PER_HOUR"},
            {"autodesk.unit.unit:tonneForceMeters-1.0.1", "TONNE_FORCE_METERS"},
            {"autodesk.unit.unit:tonneForceMetersPerMeter-1.0.1", "TONNE_FORCE_METERS_PER_METER"},
            {"autodesk.unit.unit:tonnes-1.0.1", "TONNES"},
            {"autodesk.unit.unit:tonnesForce-1.0.1", "TONNES_FORCE"},
            {"autodesk.unit.unit:tonnesForcePerMeter-1.0.1", "TONNES_FORCE_PER_METER"},
            {"autodesk.unit.unit:tonnesForcePerSquareMeter-1.0.1", "TONNES_FORCE_PER_SQUARE_METER"},
            {"autodesk.unit.unit:tonsOfRefrigeration-1.0.1", "TONS_OF_REFRIGERATION"},
            {"autodesk.unit.unit:turns-1.0.1", "TURNS"},
            {"autodesk.unit.unit:usGallons-1.0.1", "US_GALLONS"},
            {"autodesk.unit.unit:usGallonsPerHour-1.0.1", "US_GALLONS_PER_HOUR"},
            {"autodesk.unit.unit:usGallonsPerMinute-1.0.1", "US_GALLONS_PER_MINUTE"},
            {"autodesk.unit.unit:usSurveyFeet-1.0.0", "US_SURVEY_FEET"},
            {"autodesk.unit.unit:usTonnesForce-1.0.1", "US_TONNES_FORCE"},
            {"autodesk.unit.unit:usTonnesMass-1.0.1", "US_TONNES_MASS"},
            {"autodesk.unit.unit:voltAmperes-1.0.1", "VOLT_AMPERES"},
            {"autodesk.unit.unit:volts-1.0.1", "VOLTS"},
            {"autodesk.unit.unit:waterDensity4DegreesCelsius-1.0.0", "WATER_DENSITY_AT_4_DEGREES_CELSIUS"},
            {"autodesk.unit.unit:watts-1.0.1", "WATTS"},
            {"autodesk.unit.unit:wattsPerCubicFoot-1.0.1", "WATTS_PER_CUBIC_FOOT"},
            {"autodesk.unit.unit:wattsPerCubicFootPerMinute-1.0.0", "WATTS_PER_CUBIC_FOOT_PER_MINUTE"},
            {"autodesk.unit.unit:wattsPerCubicMeter-1.0.1", "WATTS_PER_CUBIC_METER"},
            {"autodesk.unit.unit:wattsPerCubicMeterPerSecond-1.0.0", "WATTS_PER_CUBIC_METER_PER_SECOND"},
            {"autodesk.unit.unit:wattsPerFoot-1.0.0", "WATTS_PER_FOOT"},
            {"autodesk.unit.unit:wattsPerMeter-1.0.0", "WATTS_PER_METER"},
            {"autodesk.unit.unit:wattsPerMeterKelvin-1.0.1", "WATTS_PER_METER_KELVIN"},
            {"autodesk.unit.unit:wattsPerSquareFoot-1.0.1", "WATTS_PER_SQUARE_FOOT"},
            {"autodesk.unit.unit:wattsPerSquareMeter-1.0.1", "WATTS_PER_SQUARE_METER"},
            {"autodesk.unit.unit:wattsPerSquareMeterKelvin-1.0.1", "WATTS_PER_SQUARE_METER_KELVIN"},
            {"autodesk.unit.unit:yards-1.0.1", "YARDS"},
            {"autodesk.unit.unit:feetFractionalInches-1.0.0", "FEET_AND_FRACTIONAL_INCHES"},
            {"autodesk.unit.unit:fractionalInches-1.0.0", "FRACTIONAL_INCHES"},
            {"autodesk.unit.unit:metersCentimeters-1.0.0", "METERS_AND_CENTIMETERS"},
            {"autodesk.unit.unit:degreesMinutes-1.0.0", "DEGREES_MINUTES_SECONDS"},
            {"autodesk.unit.unit:slopeDegrees-1.0.0", "SLOPE_DEGREES"},
            {"autodesk.unit.unit:stationingFeet-1.0.0", "FEET"},
            {"autodesk.unit.unit:stationingMeters-1.0.0", "METERS"},
            {"autodesk.unit.unit:stationingSurveyFeet-1.0.0", "US_SURVEY_FEET"},
            {"autodesk.unit.unit:ampereHours-1.0.0", "AMPERE_HOURS"},
            {"autodesk.unit.unit:ampereSeconds-1.0.0", "AMPERE_SECONDS"},
            {"autodesk.unit.unit:circularMils-1.0.0", "CIRCULAR_MILS"},
            {"autodesk.unit.unit:coulombs-1.0.0", "COULOMBS"},
            {"autodesk.unit.unit:dynes-1.0.0", "DYNES"},
            {"autodesk.unit.unit:ergs-1.0.0", "ERGS"},
            {"autodesk.unit.unit:farads-1.0.0", "FARADS"},
            {"autodesk.unit.unit:feetPerKipFoot-1.0.1", "FEET_PER_KIP_FOOT"},
            {"autodesk.unit.unit:gammas-1.0.0", "GAMMAS"},
            {"autodesk.unit.unit:gauss-1.0.0", "GAUSS"},
            {"autodesk.unit.unit:henries-1.0.0", "HENRIES"},
            {"autodesk.unit.unit:maxwells-1.0.0", "MAXWELLS"},
            {"autodesk.unit.unit:metersPerKilonewtonMeter-1.0.1", "METERS_PER_KILONEWTON_METER"},
            {"autodesk.unit.unit:mhos-1.0.0", "MHOS"},
            {"autodesk.unit.unit:microns-1.0.0", "MICRONS"},
            {"autodesk.unit.unit:mils-1.0.0", "MILS"},
            {"autodesk.unit.unit:nauticalMiles-1.0.0", "NAUTICAL_MILES"},
            {"autodesk.unit.unit:oersteds-1.0.0", "OERSTEDS"},
            {"autodesk.unit.unit:ouncesForce-1.0.0", "OUNCES_FORCE"},
            {"autodesk.unit.unit:ouncesMass-1.0.0", "OUNCES_MASS"},
            {"autodesk.unit.unit:siemens-1.0.0", "SIEMENS"},
            {"autodesk.unit.unit:slugs-1.0.0", "SLUGS"},
            {"autodesk.unit.unit:squareFeetPerKipFoot-1.0.1", "SQUARE_FEET_PER_KIP_FOOT"},
            {"autodesk.unit.unit:squareMetersPerKilonewtonMeter-1.0.1", "SQUARE_METERS_PER_KILONEWTON_METER"},
            {"autodesk.unit.unit:squareMils-1.0.0", "SQUARE_MILS"},
            {"autodesk.unit.unit:webers-1.0.0", "WEBERS"}
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
                    FamilySizeTableColumn columnHeader = selectedSizeTable.GetColumnHeader(i);

#if Revit2020 || Debug2020
                    if (versionNumber <= 2020)
                    {
                        returnedHeader.AppendFormat("{0}##{1}##{2};",
                            columnHeader.Name,
                            columnHeader.UnitType.ToString().Replace("UT_", ""),
                            columnHeader.DisplayUnitType.ToString().Replace("DUT_", ""));
                    }
#endif

#if Revit2023 || Debug2023
                    if (versionNumber >= 2021)
                    {
                        try
                        {
                            returnedHeader.AppendFormat("{0}##{1}##{2};",
                                columnHeader.Name,
                                _forgeTypeIdDict[columnHeader.GetSpecTypeId().TypeId],
                                _unitTypeIdDict[columnHeader.GetUnitTypeId().TypeId]);
                        }
                        catch
                        {
                            returnedHeader.AppendFormat("{0}##Undefined##UNDEFINED;", columnHeader.Name);
                        }
                    }
#endif
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

        private ListBox.SelectedObjectCollection SelectFromList(string title, IList<string> options)
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

        private string SelectFolder()
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                return folderBrowserDialog.ShowDialog() == DialogResult.OK ? folderBrowserDialog.SelectedPath : null;
            }
        }

        private void ConvertEncoding(string filePath, Encoding sourceEncoding, Encoding targetEncoding)
        {
            string content = File.ReadAllText(filePath, sourceEncoding);
            File.WriteAllText(filePath, content, targetEncoding);
        }
    }
}
