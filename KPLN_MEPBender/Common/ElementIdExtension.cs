using Autodesk.Revit.DB;

namespace KPLN_MEPBender.Common
{
    internal static class ElementIdExtension
    {
        public static int GetStableIntegerValue(this ElementId elementId)
        {
#if Debug2024 || Revit2024
            return (int)elementId.Value;
#else
            return elementId.IntegerValue;
#endif
        }
    }
}
