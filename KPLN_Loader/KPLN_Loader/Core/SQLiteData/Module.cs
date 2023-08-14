using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.SQLiteData
{
    internal class Module
    {
        /// <summary>
        /// Id модуля
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Id отдела
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        public int SubDepartmentId { get; set; }

        /// <summary>
        /// Путь к модулю
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// Имя модуля
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Модуль влк/выкл (True/False) для загрузки
        /// </summary>
        public string IsEnabled { get; set; }

        /// <summary>
        /// Тестовый режим вкл/выкл (True/False)
        /// </summary>
        public string IsDebugMode { get; set; }

        /// <summary>
        /// Скрытая загрузка библиотек вкл/выкл (True/False)
        /// </summary>
        public string IsLibraryModule { get; set; }
    }
}
