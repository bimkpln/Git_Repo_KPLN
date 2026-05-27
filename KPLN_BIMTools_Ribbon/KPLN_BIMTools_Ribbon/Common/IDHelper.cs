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

#if Debug2023 || Revit2023 || Debug2024 || Revit2024
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value);
#else
        internal static FilterRule CreateContainsRule(ElementId parameterId, string value) =>
            ParameterFilterRuleFactory.CreateContainsRule(parameterId, value, false);
#endif
    }
}