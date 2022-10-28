using System.Reflection;
using System.Windows;

namespace KPLN_ModelChecker_Coordinator
{
    public static class ModuleData
    {
        public static bool AutoConfirmEnabled = false;
#if Revit2020
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2018
        public static string RevitVersion = "2018";
#endif
        public static System.IntPtr MainWindowHandle { get; set; }
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        public static string Date = "2022/10/15";
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;
        public static bool ForceClose = false;
        //UserSettings
        public static bool up_send_enter = false;
        public static bool up_close_dialogs = true;
        public static bool up_notify_in_tg = false;
    }
}
