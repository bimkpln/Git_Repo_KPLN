using System.IO;
using System.Reflection;

namespace KPLN_CommandsWheel
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
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;

        /// <summary>
        /// Версия Revit, в которой запускается плагин
        /// </summary>
        public static int RevitVersion { get; set; }

        /// <summary>
        /// Полное имя текущей версии Revit, совпадающее с именем папки профиля
        /// </summary>
        public static string RevitVersionName { get; set; }

        /// <summary>
        /// Имя вкладки, в которую загрузчик добавил модуль
        /// </summary>
        public static string RibbonTabName { get; set; }

        /// <summary>
        /// Ссылка на основное окно Revit 
        /// </summary>
        public static System.IntPtr RevitMainWindowHandle { get; set; }

        private static string GetModuleFileCreationDate()
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