using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common
{
    internal static class IDHelper
    {
        internal static long EidValue(ElementId id)
        {
#if Revit2024 || Debug2024
            return id.Value;
#else
    return id.IntegerValue;
#endif
        }

        internal static int EidInt(ElementId id)
        {
#if Revit2024 || Debug2024
            return (int)id.Value;
#else
            return id.IntegerValue;
#endif
        }

    }
}
