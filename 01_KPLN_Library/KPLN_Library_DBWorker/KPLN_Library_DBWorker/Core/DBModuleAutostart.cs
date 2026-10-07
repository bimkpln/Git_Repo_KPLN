using KPLN_Library_DBWorker.Core.Abstractions;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_DBWorker.Core
{
    /// <summary>
    /// Класс файла для автостарта
    /// </summary>
    public class DBModuleAutostart : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Пользователь, которому принадлежит настройка
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int UserId { get; set; }

        /// <summary>
        /// Версия используемого Revit
        /// </summary>
        public int RevitVersion { get; set; }

        /// <summary>
        /// Пользователь, которому принадлежит настройка
        /// </summary>
        [ForeignKey(nameof(DBProject))]
        public int ProjectId { get; set; }

        /// <summary>
        /// Модуль, который должен запуститься автоматом
        /// </summary>
        [ForeignKey(nameof(DBModule))]
        public int ModuleId { get; set; }

        /// <summary>
        /// Таблица из БД, в которой я буду брать конфиг для старта
        /// </summary>
        public string DBTableName { get; set; }

        /// <summary>
        /// ID-конфигурации для старта
        /// </summary>
        public int DBTableKeyId { get; set; }

        /// <summary>
        /// Получатель уведомления об автозапуске (Users.Id).
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int? NotificationUserId { get; set; }

        /// <summary>Дата и время начала расписания, UTC.</summary>
        public DateTime? StartDateUtc { get; set; }

        /// <summary>Интервал выгрузок в часах: минимум 12, кратно 12.</summary>
        public int? IntervalHours { get; set; }

        /// <summary>Не запускать выгрузку в субботу и воскресенье по местному времени компьютера.</summary>
        public bool SkipWeekends { get; set; }

        /// <summary>Дата и время начала последней попытки выгрузки, UTC, включая попытки с ошибкой.</summary>
        public DateTime? LastStartDateUtc { get; set; }

        /// <summary>Дата и время завершения последней обработанной попытки, UTC, включая ошибки.</summary>
        public DateTime? LastExportDateUtc { get; set; }
        #endregion

        [NotMapped]
        public bool HasSchedule => StartDateUtc.HasValue && IntervalHours.HasValue
            && IntervalHours.Value >= 12 && IntervalHours.Value % 12 == 0;

        /// <summary>Старые назначения без расписания не запускаются.</summary>
        public bool IsDue(DateTime launchTimeUtc)
        {
            DateTime utc = launchTimeUtc.Kind == DateTimeKind.Local
                ? launchTimeUtc.ToUniversalTime() : DateTime.SpecifyKind(launchTimeUtc, DateTimeKind.Utc);
            DayOfWeek localDay = utc.ToLocalTime().DayOfWeek;
            if (SkipWeekends && (localDay == DayOfWeek.Saturday || localDay == DayOfWeek.Sunday))
                return false;

            launchTimeUtc = utc;
            if (!HasSchedule || launchTimeUtc < StartDateUtc.Value)
                return false;

            if (!LastStartDateUtc.HasValue || LastStartDateUtc.Value < StartDateUtc.Value)
                return true;

            if (launchTimeUtc <= LastStartDateUtc.Value)
                return false;

            // Привязываем последнюю фактическую попытку к предшествующему слоту планировщика
            // (06:00/18:00), чтобы задержка старта внутри слота не сдвигала следующий запуск.
            // StartDateUtc задаёт только фазу 12-часовой сетки, а не границы IntervalHours.
            long schedulerStepTicks = TimeSpan.FromHours(12).Ticks;
            long elapsedTicks = LastStartDateUtc.Value.Ticks - StartDateUtc.Value.Ticks;
            DateTime lastScheduledStartUtc = LastStartDateUtc.Value.AddTicks(-(elapsedTicks % schedulerStepTicks));

            // После пропуска или выходных полный интервал считаем от выполненной попытки.
            // Время завершения не влияет на расписание, очереди за пропущенные слоты нет.
            return (launchTimeUtc - lastScheduledStartUtc).TotalHours >= IntervalHours.Value;
        }

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DBEnumerator CurrentDB { get; } = DBEnumerator.ModuleAutostart;
    }
}
