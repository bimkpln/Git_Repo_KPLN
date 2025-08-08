﻿using KPLN_Loader.Common;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace KPLN_Clashes_Ribbon
{
    internal static class ModuleData
    {
        public static System.IntPtr MainWindowHandle { get; set; }
        
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
        public static int RevitVersion {  get; set; }

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
