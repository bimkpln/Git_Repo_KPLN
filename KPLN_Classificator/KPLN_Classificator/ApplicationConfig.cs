using KPLN_Classificator.ExecutableCommand;
using System;
using System.IO;
using System.Reflection;

namespace KPLN_Classificator
{
    public static class ApplicationConfig
    {
        /// <summary>
        /// Версия сборки, отображаемая в Revit
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина
        /// </summary>
        public static string Date = GetModuleCreationDate();

        /// <summary>
        /// Имя модуля
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;

        /// <summary>
        /// Флаг доступности документа. Значение false показывает, что плагин был запущен без активного объекта Document
        /// </summary>
        public static bool isDocumentAvailable { get; set; } = true;

        /// <summary>
        /// Реализация системы вывода плагина
        /// </summary>
        public static Output output;

        /// <summary>
        /// Реализация очереди для запуска команд Revit (событие OnIdling запускает команду внутри текущего UIControlledApplication)
        /// </summary>
        public static CommandEnvironment commandEnvironment;

        public static IntPtr MainWindowHandle { get; set; } //Главное окно Revit (WPF: Для определения свойства .Owner)

        private static string GetModuleCreationDate()
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
