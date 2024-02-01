using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс проекта KPLN
    /// </summary>
    public class DBProject : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Имя проекта
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Код (аббревиатура) проекта
        /// </summary>
        public string Code { get; set; }
        
        /// <summary>
        /// Стадия проектирования
        /// </summary>
        public string Stage { get; set; }

        /// <summary>
        /// Путь к корневой папке
        /// </summary>
        public string MainPath { get; set; }

        /// <summary>
        /// Путь Revit-Server
        /// </summary>
        public string RevitServerPath { get; set; }

        /// <summary>
        /// Режим блокировки проекта под набор разрешенных пользователей вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsClosed { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.Projects;
    }
}
