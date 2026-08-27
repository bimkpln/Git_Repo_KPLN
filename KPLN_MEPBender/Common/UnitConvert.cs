using Autodesk.Revit.DB;

namespace KPLN_MEPBender.Common
{
    internal static class UnitConvert
    {
        public static double MmToInternal(double valueMm)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertToInternalUnits(valueMm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(valueMm, UnitTypeId.Millimeters);
#endif
        }
    }
}
