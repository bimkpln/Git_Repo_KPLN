using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Scoper
{
    public static class ModuleData
    {
        public static Guid SessionGuid = Guid.NewGuid();
        public static bool IsDebugMode = KPLN_Loader.Preferences.User.SystemName == "iperfilyev";
    }
}
