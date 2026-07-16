using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace KPLN_Tools
{
    internal sealed class WindowHandleWrapper : IWin32Window
    {
        public WindowHandleWrapper(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }
    }

    internal static class ModuleData
    {
        public static IntPtr MainWindowHandle { get; set; }

        internal static IWin32Window MainWindowOwner => new WindowHandleWrapper(MainWindowHandle);

        /// <summary>
        /// Версия сборки, отображаемая в Revit
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина
        /// </summary>
        public static string Date = GetModuleCreationDate();

        /// <summary>
        /// Версия Revit, в которой запускается плагин
        /// </summary>
        public static int RevitVersion { get; set; }

        /// <summary>
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;
        public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();

        private static string GetModuleCreationDate()
        {
            string filePath = Assembly.GetExecutingAssembly().Location;
            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                return fileInfo.LastWriteTime.ToString("yyyy/MM/dd");
            }

            return "Дата не определена";
        }
    }
}
