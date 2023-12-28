using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс проекта KPLN
    /// </summary>
    public class DBProject : IDBEntity
    {
        [Key]
        public int Id { get; set; }

        public DB_Enumerator CurrentDB { get; set; }

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
        /// Режим блокировки проекта под набор разрешенных пользователей вкл/выкл (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsClosed { get; set; }
    }
}
