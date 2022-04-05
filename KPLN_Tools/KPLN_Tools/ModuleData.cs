﻿using System.Windows;

namespace KPLN_Tools
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
        public static System.IntPtr MainWindowHandle { get; set; }//Главное окно Revit (WPF: Для определения свойства .Owner)
        public static string Build = string.Format("Revit {0}", RevitVersion);
        public static string Version = "1.0.0.2";//Отображаемая версия модуля в Revit
        public static string Date = "2022/04/05";//Дата последнего изменения
        public static string ManualPage = "https://kpln.kdb24.ru/article/60264/";//Ссылка на web-страницу по клавише F1
        public static string ModuleName = "Tools";//Имя модуля

    }
}
