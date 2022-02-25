using System.Windows;

namespace KPLN_RevitWorksets
{
    public static class ModuleData
    {
#if Revit2022
        public static string RevitVersion = "2022";
        public static Window RevitWindow { get; set; }
#endif
#if R2020
        public static string RevitVersion = "2020";
        public static Window RevitWindow { get; set; }
#endif
#if R2018
        public static string RevitVersion = "2018";
#endif
        public static System.IntPtr MainWindowHandle { get; set; }//Главное окно Revit (WPF: Для определения свойства .Owner)
        public static string Build = string.Format("Revit {0}", RevitVersion);
        public static string Version = "1.0.0.1";//Отображаемая версия модуля в Revit
        public static string Date = "2022/02/25";//Дата последнего изменения
        public static string ManualPage = "http://moodle.stinproject.local/mod/book/view.php?id=396&chapterid=437";//Ссылка на web-страницу по клавише F1
        public static string ModuleName = "RevitWorksets";//Имя модуля
    }
}
