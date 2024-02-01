using KPLN_Library_SQLiteWorker.Core.SQLiteData.Abstractions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPLN_Library_SQLiteWorker.Core.SQLiteData
{
    /// <summary>
    /// Класс файла для обмена
    /// </summary>
    public class DBRevitDocExchanges : IDBEntity
    {
        #region Столбцы из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Проект, к которому файл относится
        /// </summary>
        [ForeignKey(nameof(DBProject))]
        public int ProjectId { get; set; }

        /// <summary>
        /// Связь с документом из БД
        /// </summary>
        [ForeignKey(nameof(DBDocument))]
        public int DocumentId { get; set; }

        /// <summary>
        /// Путь к файлу/папке откуда копируем
        /// </summary>
        public string PathFrom { get; set; }

        /// <summary>
        /// Путь к корневой папке куда копируем
        /// </summary>
        public string PathTo { get; set; }

        /// <summary>
        /// Режим блокировки файла под действия по обмену (True/False). В БД тип данных текст, преобразование происходит в Dapper
        /// </summary>
        public bool IsActive { get; set; }
        #endregion

        /// <summary>
        /// Привязка к БД из DB_Enumerator
        /// </summary>
        public static DB_Enumerator CurrentDB { get; } = DB_Enumerator.RevitDocExchanges;
    }
}
