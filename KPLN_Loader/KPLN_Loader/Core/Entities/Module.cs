using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Loader.Core.Entities
{
    internal sealed class Module
    {
        /// <summary>
        /// Id модуля
        /// </summary>
        [Key]
        internal int Id { get; set; }

        /// <summary>
        /// Id отдела
        /// </summary>
        [ForeignKey(nameof(SubDepartment))]
        internal int SubDepartmentId { get; set; }

        /// <summary>
        /// Путь к модулю
        /// </summary>
        internal string Path { get; set; }

        /// <summary>
        /// Имя модуля
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// Модуль влк/выкл (True/False) для загрузки. В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        internal bool IsEnabled { get; set; }

        /// <summary>
        /// Тестовый режим вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper. Для модулей библиотек - он всегда False
        /// </summary>
        internal bool IsDebugMode { get; set; }

        /// <summary>
        /// Скрытая загрузка библиотек вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        internal bool IsLibraryModule { get; set; }
    }
}
