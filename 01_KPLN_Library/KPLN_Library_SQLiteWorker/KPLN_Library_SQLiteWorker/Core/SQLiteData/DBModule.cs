using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс модуля KPLN
    /// </summary>
    public class DBModule : IDBEntity
    {
        [Key]
        public int Id { get; set; }

        public DB_Enumerator CurrentDB { get; set; }

        /// <summary>
        /// Id отдела
        /// </summary>
        [ForeignKey(nameof(DBSubDepartment))]
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
        /// Модуль влк/выкл (True/False) для загрузки. В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Тестовый режим вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper. Для модулей библиотек - он всегда False
        /// </summary>
        public bool IsDebugMode { get; set; }

        /// <summary>
        /// Скрытая загрузка библиотек вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsLibraryModule { get; set; }
    }
}
