using Autodesk.Revit.DB;

namespace KPLN_Tools_AR.Common
{
    internal static class IDHelper
    {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
        internal static long ElIdValue(ElementId id) => id.IntegerValue;
#else
        internal static long ElIdValue(ElementId id) => id.Value;
#endif

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
        internal static int ElIdInt(ElementId id) => id.IntegerValue;
#else
        internal static int ElIdInt(ElementId id) => checked((int)id.Value);
#endif
        internal static ElementId CreateElementId(long value)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return new ElementId(checked((int)value));
#else
            return new ElementId(value);
#endif
        }

#if Debug2020 || Revit2020
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value, false);
#else
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value);
#endif

        internal static double ConvertMmToInternal(int valueMm)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#endif
        }

        internal static double ConvertInternalToMm(double valueInternal)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#endif
        }

        internal static double ConvertInternalAreaToSquareMeters(double valueInternal)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_SQUARE_METERS);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.SquareMeters);
#endif
        }

        internal static bool IsNumber(Parameter parameter)
        {
#if Debug2020 || Revit2020
            return parameter.Definition.ParameterType == ParameterType.Number;
#else
            return parameter.Definition.GetDataType() == SpecTypeId.Number;
#endif
        }

        internal static bool IsLength(Parameter parameter)
        {
#if Debug2020 || Revit2020
            return parameter.Definition.ParameterType == ParameterType.Length;
#else
            return parameter.Definition.GetDataType() == SpecTypeId.Length;
#endif
        }

        internal static bool IsAngle(Parameter parameter)
        {
#if Debug2020 || Revit2020
            return parameter.Definition.ParameterType == ParameterType.Angle;
#else
            return parameter.Definition.GetDataType() == SpecTypeId.Angle;
#endif
        }
    }
}