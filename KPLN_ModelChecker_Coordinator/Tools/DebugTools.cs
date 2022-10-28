using System;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_Coordinator.Tools
{
    public static class DebugTools
    {
        public static void Error(Exception e)
        {
            if (e.InnerException != null) Error(e.InnerException);
            PrintError(e);
        }
    }
}
