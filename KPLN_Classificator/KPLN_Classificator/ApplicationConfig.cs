using KPLN_Classificator.ExecutableCommand;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

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
        /// Реализация способа получения данных о пользователе плагина
        /// </summary>
        public static UserInfo.UserInfo userInfo;
        /// <summary>
        /// Реализация очереди для запуска команд Revit (событие OnIdling запускает команду внутри текущего UIControlledApplication)
        /// </summary>
        public static CommandEnvironment commandEnvironment;
#if Revit2020 || Revit2020Std
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if Revit2018 || Revit2018Std
        public static string RevitVersion = "2018";
#endif
        public static IntPtr MainWindowHandle { get; set; } //Главное окно Revit (WPF: Для определения свойства .Owner)
    }
}
