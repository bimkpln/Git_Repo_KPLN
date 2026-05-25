using KPLN_Library_DBWorker.Core.Abstractions;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_DBWorker.Core
{
    /// <summary>
    /// Класс активности по плагину
    /// </summary>
    public class DBPluginActivity : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Ссылка на модуль
        /// </summary>
        [ForeignKey(nameof(DBModule))]
        public int ModuleId { get; set; }

        /// <summary>
        /// Имя плагина, который запустился с модуля
        /// </summary>
        public string PluginName { get; set; }

        /// <summary>
        /// Ссылка на отдел
        /// </summary>
        [Obsolete("25.05.26 - удалить через пару месяцев. Используем UserId вместо этого, чтобы не дублировать инфу в БД. Сразу заменить не вариант - сломает таблицу")]
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }

        /// <summary>
        /// Ссылка на пользователя
        /// </summary>
        [ForeignKey(nameof(DBUser))]
        public int UserId { get; set; }

        /// <summary>
        /// Колчиество запусков
        /// </summary>
        public int UsageCount { get; set; }

        /// <summary>
        /// Дата последнего запуска
        /// </summary>
        public string LastActivityDate { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DBEnumerator CurrentDB { get; } = DBEnumerator.PluginsActivity;
    }
}
