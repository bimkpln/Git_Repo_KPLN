using Autodesk.Revit.DB;

namespace KPLN_Tools.Common
{
    internal static class IDHelper
    {
#if !Revit2024 && !Debug2024
        internal static int ElIdValue(ElementId id) => id.IntegerValue;
#else
        internal static long ElIdValue(ElementId id) => id.Value;
#endif

#if !Revit2024 && !Debug2024
        internal static int ElIdInt(ElementId id) => id.IntegerValue;
#else
        internal static int ElIdInt(ElementId id) => (int)id.Value;
#endif
    }
}
