using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_ModelChecker_User
{
    public static class ModuleData
    {
#if Revit2020
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2018
        public static string RevitVersion = "2018";
#endif
        public static System.IntPtr MainWindowHandle { get; set; }
        public static string Build = string.Format("Revit {0}", RevitVersion);
        public static string Version = "1.0.0.2b";
        public static string Date = "2020/10/02";
        public static string ManualPage = "https://kpln.kdb24.ru/article/60264/";
        public static string ModuleName = "Model checker";
        public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();
    }
}
