using KPLN_Loader.Common;
using KPLN_Publication.Forms;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;

namespace KPLN_Publication
{
    public static class ModuleData
    {
#if Revit2023
        public static string RevitVersion = "2023";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2022
        public static string RevitVersion = "2022";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2020
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2018
        public static string RevitVersion = "2018";
#endif
        public static System.IntPtr MainWindowHandle { get; set; }
        /// <summary>
        /// Версия сборки, отображаемая в Revit
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина
        /// </summary>
        public static string Date = "2023/06/01";

        /// <summary>
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;
        //public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();
        public static SetManager Form { get; set; }
    }
}
