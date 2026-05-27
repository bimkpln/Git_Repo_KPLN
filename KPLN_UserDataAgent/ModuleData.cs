using System;
using System.IO;
using System.Reflection;

namespace KPLN_UserDataAgent
{
    /// <summary>
    /// Доплнительные атрибуты по текущему модулю для отображения в Revit.
    /// </summary>
    internal static class ModuleData
    {
        /// <summary>
        /// Версия сборки.
        /// </summary>
        public static string Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Актуальная дата плагина.
        /// </summary>
        public static string Date = GetModuleFileCreationDate();

        /// <summary>
        /// Имя модуля.
        /// </summary>
        public static string ModuleName = Assembly.GetExecutingAssembly().GetName().Name;

        /// <summary>
        /// Версия Revit, в которой запускается плагин.
        /// </summary>
        public static int RevitVersion { get; set; }

        /// <summary>
        /// Ссылка на основное окно Revit.
        /// </summary>
        public static IntPtr RevitMainWindowHandle { get; set; }

        /// <summary>
        /// Путь к общей SQLite-базе агента.
        /// </summary>
        public const string CentralDatabasePath = @"Z:\Отдел BIM\Туленинов Роман\СТАТИСТИКА\KPLN_UserDataAgent.db";

        /// <summary>
        /// Compatibility alias for older module code.
        /// </summary>
        public const string DatabasePath = CentralDatabasePath;

        /// <summary>
        /// Задержка первой попытки отправки локальной очереди в общую БД.
        /// </summary>
        public const int SyncStartDelaySeconds = 10;

        /// <summary>
        /// Интервал попыток отправки локальной очереди в общую БД.
        /// </summary>
        public const int SyncIntervalSeconds = 60;

        /// <summary>
        /// Задержка ускоренной попытки синхронизации после записи события.
        /// </summary>
        public const int SyncAfterWriteDelaySeconds = 2;

        /// <summary>
        /// Размер пачки событий за одну попытку отправки.
        /// </summary>
        public const int SyncBatchSize = 500;

        /// <summary>
        /// Таймаут ожидания локального SQLite lock.
        /// </summary>
        public const int LocalBusyTimeoutMs = 1000;

        /// <summary>
        /// Таймаут ожидания сетевого SQLite lock.
        /// </summary>
        public const int CentralBusyTimeoutMs = 1000;

        /// <summary>
        /// Минимальный интервал между debug-диалогами.
        /// </summary>
        public const int DebugDialogMinIntervalSeconds = 300;

        /// <summary>
        /// Временный активный режим отладки: ошибки показываются пользователю.
        /// Для тихого режима заменить на false.
        /// </summary>
        public const bool ShowDebugErrors = true;

        /// <summary>
        /// Локальная SQLite-база-очередь на диске пользователя.
        /// </summary>
        public static string LocalDatabasePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KPLN",
                    "UserDataAgent",
                    "KPLN_UserDataAgent_Local.db");
            }
        }

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