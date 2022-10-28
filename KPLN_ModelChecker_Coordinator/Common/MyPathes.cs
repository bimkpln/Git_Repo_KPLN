namespace KPLN_ModelChecker_Coordinator.Common
{
    /// <summary>
    /// Класс-обертка для централизованного управления ссылками на сервер
    /// </summary>
    internal static class MyPathes
    {
        /// <summary>
        /// Путь к основной базе данных
        /// </summary>
        public static readonly string MainDBConnection = KPLN_Library_DataBase.DbControll.MainDBConnection;

        /// <summary>
        /// Путь к базе данных под модуль проверок
        /// </summary>
        public static readonly string ModelCheckerDBPath = @"Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader\DB\KPLN_ModelChecker_Coordinator";

        /// <summary>
        /// Путь к папке со скринами под модуль проверок
        /// </summary>
        public static readonly string ModelCheckerCommonPath = @"Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader\DB\TaskDialogs";
    }
}
