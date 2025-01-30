using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
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
        [ForeignKey(nameof(DBSubDepartment))]
        public int SubDepartmentId { get; set; }

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
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.PluginsActivity;
    }
}
