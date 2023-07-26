using KPLN_Loader.Common;
using KPLN_Publication.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace KPLN_Publication
{
    public static class ModuleData
    {
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
        public static string Build = string.Format("built for Revit {0}", RevitVersion);
        public static string Version = "1.2.0.0";
        public static string Date = "2023/07/25";
        public static string ManualPage = "https://kpln.kdb24.ru/article/89149/";
        public static string ModuleName = "Publishing";
        public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();
        public static SetManager Form { get; set; }
    }
}
