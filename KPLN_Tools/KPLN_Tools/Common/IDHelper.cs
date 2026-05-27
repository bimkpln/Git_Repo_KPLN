using Autodesk.Revit.DB;

namespace KPLN_Tools.Common
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
        internal static int ElIdInt(ElementId id) => (int)id.Value;
#endif
        internal static ElementId CreateElementId(long value)
        {
#if Revit2024 || Debug2024
            return new ElementId(value);
#else
            return new ElementId((int)value);
#endif
        }

#if Debug2023 || Revit2023 || Debug2024 || Revit2024
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value);
#else
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value, false);
#endif

        internal static double ConvertMmToInternal(int valueMm)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        internal static double ConvertInternalToMm(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.Millimeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_MILLIMETERS);
#endif
        }

        internal static double ConvertInternalAreaToSquareMeters(double valueInternal)
        {
#if Revit2024 || Revit2023 || Debug2024 || Debug2023
            return UnitUtils.ConvertFromInternalUnits(valueInternal, UnitTypeId.SquareMeters);
#else
            return UnitUtils.ConvertFromInternalUnits(valueInternal, DisplayUnitType.DUT_SQUARE_METERS);
#endif
        }
    }
}