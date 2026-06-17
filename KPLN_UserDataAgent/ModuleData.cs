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
        /// Ссылка на основное окно Revit.
        /// </summary>
        public static IntPtr RevitMainWindowHandle { get; set; }

        /// <summary>
        /// Корневая папка общей статистики.
        /// </summary>
        public const string StatisticsDirectory = @"Z:\Отдел BIM\Туленинов Роман\СТАТИСТИКА";

        /// <summary>
        /// Корневая папка центральных SQLite-баз агента.
        /// Внутри создаются файлы по отделу и месяцу с UserEvents, EventTransactions и AgentErrors.
        /// </summary>
        public static string CentralDatabasePath
        {
            get { return Path.Combine(StatisticsDirectory, "UserEvents"); }
        }

        public static string DatabasePath
        {
            get { return CentralDatabasePath; }
        }

        /// <summary>
        /// Справочная SQLite-база пользователей KPLN Loader.
        /// Используется для определения отдела пользователя при выборе центральной БД.
        /// </summary>
        public const string ReferenceDatabasePath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_MainDB.db";

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

        /// <summary>
        /// Локальный кэш последнего успешно прочитанного соответствия пользователей и отделов.
        /// Если справочная БД временно недоступна, агент использует этот кэш.
        /// </summary>
        public static string LocalDepartmentLookupCachePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KPLN",
                    "UserDataAgent",
                    "KPLN_UserDataAgent_DepartmentLookup.cache");
            }
        }

        /// <summary>
        /// Задержка первой фоновой попытки отправить локальную очередь в общую БД после запуска Revit.
        /// Единица измерения: секунды.
        /// </summary>
        public const int SyncStartDelaySeconds = 60;

        /// <summary>
        /// Интервал регулярных фоновых попыток отправить локальную очередь в общую БД.
        /// Если общая БД или сетевой диск недоступны, следующая попытка будет через этот интервал.
        /// Единица измерения: секунды.
        /// </summary>
        public const int SyncIntervalSeconds = 300;

        /// <summary>
        /// Задержка ускоренной попытки синхронизации после записи нового события в локальную БД.
        /// Единица измерения: секунды.
        /// </summary>
        public const int SyncAfterWriteDelaySeconds = 60;

        /// <summary>
        /// Случайная добавка к задержкам синхронизации, чтобы пользователи не били в общую БД одновременно.
        /// Фактическая задержка = базовая задержка + случайное число от 0 до этого значения.
        /// Единица измерения: секунды.
        /// </summary>
        public const int SyncRandomJitterSeconds = 120;

        /// <summary>
        /// Максимальное количество локальных событий, отправляемых в общую БД за одну попытку синхронизации.
        /// Единица измерения: штуки записей.
        /// </summary>
        public const int SyncBatchSize = 200;

        /// <summary>
        /// Количество последних месяцев, которые хранятся в центральных базах агента.
        /// Текущий месяц входит в этот лимит.
        /// </summary>
        public const int CentralDatabaseRetentionMonths = 4;

        /// <summary>
        /// Таймаут ожидания блокировки локальной SQLite-базы пользователя.
        /// Единица измерения: миллисекунды.
        /// </summary>
        public const int LocalBusyTimeoutMs = 1000;

        /// <summary>
        /// Таймаут ожидания блокировки общей SQLite-базы на сетевом диске.
        /// Единица измерения: миллисекунды.
        /// </summary>
        public const int CentralBusyTimeoutMs = 5000;

        /// <summary>
        /// Максимальный размер текущей локальной SQLite-базы перед ротацией в архивную очередь.
        /// Архивная очередь продолжает синхронизироваться в общую БД и удаляется только после успешной отправки всех записей.
        /// Единица измерения: мегабайты.
        /// </summary>
        public const int LocalDatabaseRotationSizeMb = 50;


        /// <summary>
        /// Режим показа ошибок пользователю через диалог Revit.
        /// false: скрытый режим, ошибки не показываются пользователю, но пишутся в техническую очередь AgentErrors.
        /// </summary>
        public const bool ShowDebugErrors = false;

        /// <summary>
        /// Минимальный интервал между debug-диалогами с ошибками.
        /// Единица измерения: секунды.
        /// </summary>
        public const int DebugDialogMinIntervalSeconds = 300;


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
