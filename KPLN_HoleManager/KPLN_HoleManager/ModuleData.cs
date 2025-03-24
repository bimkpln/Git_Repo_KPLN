using System.IO;
using System.Reflection;

namespace KPLN_HoleManager
{
    /// <summary>
    /// Доплнительные атрибуты по текущему модулю для отображения в Revit
    /// </summary>
    internal static class ModuleData
    {
        /// Версия сборки
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// Актуальная дата плагина
        public static string Date = GetModuleFileCreationDate();

        /// Имя модуля
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
