using KPLN_Classificator.ExecutableCommand;
using System;

namespace KPLN_Classificator
{
    public static class ApplicationConfig
    {
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
    }
}
