using System.IO;
using System.Reflection;

namespace KPLN_BIMTools_Ribbon
{
    /// <summary>
    /// Доплнительные атрибуты по текущему модулю для отображения в Revit
    /// </summary>
    internal static class ModuleData
    {
        /// <summary>
        /// Версия сборки
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина
        /// </summary>
        public static string Date = GetModuleFileCreationDate();

        /// <summary>
        /// Версия Revit, в которой запускается плагин
        /// </summary>
        public static int RevitVersion { get; set; }

        /// <summary>
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;

        private static string GetModuleFileCreationDate()
        {
            string filePath = Assembly.GetExecutingAssembly().Location;
            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                return fileInfo.CreationTime.ToString("yyyy/MM/dd");
            }

            return "Дата не определена";
        }
    }
}
