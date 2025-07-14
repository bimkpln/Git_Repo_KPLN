using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
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
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.ModuleAutostart;
    }
}
